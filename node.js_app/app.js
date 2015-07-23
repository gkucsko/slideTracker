/**
 * Module dependencies.
 */

var express = require('express');
var app = express();

// filesystem
var fs = require('fs');

var privateKey  = fs.readFileSync('../SSL_cert/slidetracker.key', 'utf8');
var certificate = fs.readFileSync('../SSL_cert/slidetracker.crt', 'utf8');
// var bundle = fs.readFileSync('../SSL_cert/gd_bundle.crt', 'utf8');
var credentials = {key: privateKey, cert: certificate};
// var credentials = {key: privateKey, cert: certificate, ca: [bundle]};

var https = require('https').Server(credentials,app);
var io = require('socket.io')(https);
var path = require('path');

// mongoDB abstraction layer
var mongoose = require('mongoose');

// log requests to the console (express4)
var morgan = require('morgan');

// pull information from HTML POST (express4)
var bodyParser = require('body-parser');

// simulate DELETE and PUT (express4)
var methodOverride = require('method-override');

// file uploads
var multer = require('multer');

// cron job
var cron = require('cron');
var CronJob = cron.CronJob;

// json to csv
var json2csv = require('json2csv');

// amazon S3 storage
//var AWS = require('aws-sdk');
//AWS.config.update({region: 'us-east-1'});

// delete directories and files
var rimraf = require('rimraf');

// static key used to avoid bots posting
var apiKey = 'N3sN7AiWTFK9XNwSCn7um35joV6OFslL';

app.set('port', 3000);
app.use(express.static(path.join(__dirname, 'public')));

// create a write stream (in append mode)
var accessLogStream = fs.createWriteStream(__dirname + '/access.log', {flags: 'a'})

// setup the logger
app.use(morgan('combined', {stream: accessLogStream}))

// log every request to the console
app.use(bodyParser.urlencoded({
	'extended' : 'true'
}));

// parse application/x-www-form-urlencoded
app.use(bodyParser.json());

// parse application/json
app.use(bodyParser.json({
	type : 'application/vnd.api+json'
}));

// parse application/vnd.api+json as json
app.use(methodOverride());
app.use(multer({
	dest : './uploads/'
}));
app.use(multer({
	dest : './uploads/',
	rename : function(fieldname, filename) {
		return filename.replace(/\W+/g, '-').toLowerCase() + Date.now()
	},
	limits : {
		fileSize : 20000000, // max 20 MB
		files : 1
	}
}))

// database connection
// use environment variable MONGODB_STRING to store connection string 
mongoose.connect('mongodb://'+process.env.MONGODB_STRING);

// Authenticator for admin section
app.use(function(req, res, next) {
    var auth;
	if (req.url.indexOf('/admin/') === 0) // 'hidden' Admin Section
	{
	    // check whether an autorization header was send    
	    if (req.headers.authorization) {
	      // only accepting basic auth, so:
	      // * cut the starting "Basic " from the header
	      // * decode the base64 encoded username:password
	      auth = new Buffer(req.headers.authorization.substring(6), 'base64').toString().split(':');
	    }
	    // checks password
	    if (!auth || auth[0] !== 'admin' || auth[1] !== process.env.ADMIN_PW) {
	        res.statusCode = 401;
	        res.setHeader('WWW-Authenticate', 'Basic realm="AdminSection"');
	        res.end('Unauthorized');
	        return;
	    }
	}
	next();
});

var db = mongoose.connection;
db.on('error', console.error.bind(console, 'connection error:'));
db.once('open', function(callback) {
	console.log('mongodb connected');
});

// define main presentation scheme
var presSchema = mongoose.Schema({
	pres_ID : String,
	creator : String,
	n_slides : Number,
	cur_slide : Number,
	file_size : Number,
	clients : Number,
	active : Boolean,
	download : Boolean,
	created : Date,
	updated : Date
})

presSchema.methods.toJSON = function() {
  var obj = this.toObject()
  delete obj._id
  return obj
}

var Presentation = mongoose.model('Presentation', presSchema)

// define log scheme
var logSchema = mongoose.Schema({
	unique_ID: String,
	creator : String,
	n_slides : Number,
	file_size : Number,
	clients : Number,
	active : Boolean,
	download : Boolean,
	created : Date
})

var LogPres = mongoose.model('LogPres', logSchema)

// present favicon
var favicon = require('serve-favicon');
app.use(favicon(path.join(__dirname,'public','images','favicon.ico')));

//define delete CRON job
new CronJob('00 00 * * *', function(){
	//find presentations older than 1 day
	var cutoff = new Date();
	cutoff.setDate(cutoff.getDate()-1);
	Presentation.find({updated: {$lt: cutoff}})
	.exec(function(err, pres) {
			if (err) {
				res.status(400).json(err);
				return;
			}
			if (pres[0]) {				
				for (i = 0; i < pres.length; i++) {
					
					//console.log('deleted: '+pres[i].pres_ID);
					
					// delete database entry
					Presentation.find({ 'pres_ID' : pres[i].pres_ID }).remove().exec();
			
					// delete folder and files
					rimraf('./public/files/' + pres[i].pres_ID, function(req, res) { });   				
				}				
			}
	});
}, null, true, "America/New_York");

// create presentation
app.post('/api/v1/presentations', function(req, res) {

	// check inputs	
	if (!req.body.key || !req.body.creator || !req.body.n_slides) {
		var err = 'sorry, problem with input parameters';
		res.status(400).json(err);
		return;
	}
	
	// check key
	if (req.body.key != apiKey) {
		var err = 'sorry, you are not authorized';
		res.status(401).json(err);
		return;
	}
	var response = {};
	
	// generate 3 random letters and a random integer 10-99
 	var pres_ID = "";
    var possible = "abcdefghijkmnpqrstuvwxyz";
    for( var i=0; i < 3; i++ ){pres_ID += possible.charAt(Math.floor(Math.random() * possible.length));}
	pres_ID += Math.floor((Math.random() * 89) + 10);
	
	//check if pres_ID already taken
	Presentation.find({ 'pres_ID' : pres_ID })
	.exec(function(err, pres) {
		if (err) {
			res.status(400).json(err);
			return;
		}
		if (pres[0]) {
			res.status(400).json('could not find a unique ID for you');
			return;
		}
		
		//var pres_ID = req.body.pres_ID;
		var creator = req.body.creator;
		var n_slides = req.body.n_slides;
		var now = new Date();
		var created = now.toJSON();

		//create db entry
		var new_pres = new Presentation({
			pres_ID : pres_ID,
			creator : creator,
			n_slides : n_slides,
			cur_slide : 1,
			file_size : 0,
			clients : 0,
			active : false,
			download : false,
			created : created,
			updated : created
		});

		//save db entry
		new_pres.save(function(err) {
			if (err) {
				res.status(400).json(err);
				return;
			}else{
				fs.mkdir('./public/files/' + pres_ID);
				var passHash = new_pres._id;		
				res.status(201).json({ presentation: new_pres , passHash: passHash });
			}
		});

		//create log entry
		var new_log = new LogPres({
			unique_ID : pres_ID+'_'+created,
			creator : creator,
			n_slides : n_slides,
			file_size : 0,
			clients : 0,
			active : false,
			download : false,
			created : created
		});

		//save log db entry
		new_log.save(function(err) {});
		
	});
});

// get presentation info
app.get('/api/v1/presentations/:pres_ID', function(req, res) {
	// remove special characters and whitespace
	var lookup_ID = req.params.pres_ID.replace(/[^\w\s]/gi, '');
	lookup_ID = lookup_ID.replace(/ /g,'');
	// make lowercase
	lookup_ID = lookup_ID.toLowerCase();
	// get presentation
	Presentation.find({ 'pres_ID' : lookup_ID })
	.exec(function(err, pres) {
		if (err) {
			res.status(400).json(err);
			return;
		}
		res.status(200).json(pres);
	});
});

// update presentation
app.put('/api/v1/presentations/:pres_ID', function(req, res) {
	
	// check inputs	
	if (!req.body.key || !req.body.passHash || !req.body.cur_slide || !req.body.n_slides || !req.body.active) {
		var err = 'incorrect input';
		res.status(400).json(err);
		return;
	}
	
	// check key
	if (req.body.key != apiKey) {
		var err = 'sorry, you are not authorized';
		res.status(401).json(err);
		return;
	}
	
	Presentation.find({ 'pres_ID' : req.params.pres_ID })
	.exec(function(err, pres) {
		if (err) {
			res.status(400).json(err);
			return;
		}
		if (!pres[0]) {
			res.status(400).json('did not find presention');
			return;
		}
		var r_pres = pres[0];
		
		// check password
		if (req.body.passHash != r_pres._id) {
			var err = 'sorry, you need to verify this is your presentation';
			res.status(401).json(err);
			return;
		}		

		// update log entry if made active
		if (r_pres.active==false && req.body.active==true) {
			LogPres.find({ 'unique_ID' : r_pres.pres_ID+'_'+r_pres.created.toISOString()})
			.exec(function(err, lpres) {
				var log_pres = lpres[0];
				log_pres.file_size = r_pres.file_size;
				log_pres.active = true;
				log_pres.download = r_pres.download;
				log_pres.save(function(err) {})
			});
		}
		
		r_pres.cur_slide = req.body.cur_slide;
		r_pres.n_slides = req.body.n_slides;
		r_pres.active = req.body.active;
		var now = new Date();
		var updated = now.toJSON();
		r_pres.updated = updated;
		r_pres.save(function(err) {
			if (err) {
				res.status(400).json(err);
				return;
			}
			res.status(200).json({ presentation: r_pres });
		});
		if (r_pres.active) {
			io.emit('update', req.params.pres_ID);
		}
	});
});

// delete presentation
app.put('/api/v1/presentations/:pres_ID/delete', function(req, res) {
	
	if (!req.body.key || !req.body.passHash) {
		var err = 'sorry, problem with input parameters';
		res.status(400).json(err);
		return;
	}
	
	// check key
	if (req.body.key != apiKey) {
		var err = 'sorry, you are not authorized';
		res.status(401).json(err);
		return;
	}

	Presentation.find({ 'pres_ID' : req.params.pres_ID })
	.exec(function(err, pres) {
		if (err) {
			res.status(400).json(err);
			return;
		}
		if (!pres[0]) {
			res.status(200).json('didnt find presentation, so no deletion needed');;
			return;
		}
		var r_pres = pres[0];
		
		// check password
		if (req.body.passHash != r_pres._id) {
			var err = 'sorry, you need to verify this is your presentation';
			res.status(401).json(err);
			return;
		}else{
			// send finish command to clients
			io.emit('quit', req.params.pres_ID);
			
			// delete database entry
			Presentation.find({ 'pres_ID' : req.params.pres_ID }).remove().exec();
			
			// delete folder and files
			rimraf('./public/files/' + req.params.pres_ID, function(req, res) { });
			res.status(200).json('deleted');
		}	
	});
});

// handle slide uploads
app.post('/api/v1/presentations/:pres_ID/slides', function(req, res) {

	// test inputs
	if (!req.body.key || !req.body.passHash || !req.params.pres_ID || !req.body.slide_ID || !req.files) {
		var err = 'sorry, problem with input parameters';
		res.status(400).json(err);
		return;
	}

	// check key
	if (req.body.key != apiKey) {
		var err = 'sorry, you are not authorized';
		res.status(401).json(err);
		return;
	}
	var response = {};

	//check if correct filetype
	if (req.files.slide.mimetype != 'image/png' || req.files.slide.extension != 'PNG') {
		var err = 'sorry, problem with file extension';
		res.status(400).json(err);
		return;
	}
	var pres_ID = req.params.pres_ID;
	var slide_ID = req.body.slide_ID;
	var old_filename = req.files.slide.name;
	var filesize = req.files.slide.size;
	
	Presentation.find({ 'pres_ID' : req.params.pres_ID })
	.exec(function(err, pres) {
		if (err) {
			res.status(400).json(err);
			return;
		}		
		if (!pres[0]) {
			res.status(400).json('did not find presention');
			return;
		}
		var r_pres = pres[0];
		
		// check password
		if (req.body.passHash != r_pres._id) {
			var err = 'sorry, you need to verify this is your presentation';
			res.status(401).json(err);
			return;
		}else{
			//check if filesize too big
			if (filesize+r_pres.file_size > 20000000) {
				var err = 'sorry, total filesize too large';
				res.status(400).json(err);
				return;
			}
			//check if slide file already exists
			if (fs.existsSync('./public/files/' + pres_ID + '/Slide' + slide_ID + '.PNG')) {
				var err = 'sorry, slide already exists';
				res.status(400).json(err);
				return;
			}
			r_pres.file_size = r_pres.file_size + filesize;
			r_pres.save(function(err) {
			if (err) {
				res.status(400).json(err);
				return;
			}
				fs.rename('./uploads/' + old_filename, './public/files/' + pres_ID + '/Slide' + slide_ID + '.PNG');
				res.status(201).json('upload succeeded!');
			});
		}
	});
});

// handle PDF uploads
app.post('/api/v1/presentations/:pres_ID/presentation', function(req, res) {

	// test inputs
	if (!req.body.key || !req.body.passHash || !req.params.pres_ID || !req.files) {
		var err = 'sorry, problem with input parameters';
		res.status(400).json(err);
		return;
	}
	
	// check key
	if (req.body.key != apiKey) {
		var err = 'sorry, you are not authorized';
		res.status(401).json(err);
		return;
	}
	var response = {};
	
	//check if correct filetype
	if (req.files.pres.mimetype != 'application/pdf' || req.files.pres.extension != 'pdf') {
		var err = 'sorry, problem with file extension';
		res.status(400).json(err);
		return;
	}
	var pres_ID = req.params.pres_ID;
	var old_filename = req.files.pres.name;
	
	Presentation.find({ 'pres_ID' : req.params.pres_ID })
	.exec(function(err, pres) {
		if (err) {
			res.status(400).json(err);
			return;
		}
		if (!pres[0]) {
			res.status(400).json('did not find presention');
			return;
		}
		var r_pres = pres[0];
		
		// check password
		if (req.body.passHash != r_pres._id) {
			var err = 'sorry, you need to verify this is your presentation';
			res.status(401).json(err);
			return;
		}else{
	
			//check if slide file already exists
			if (fs.existsSync('./public/files/' + pres_ID + '/presentation.pdf')) {
				var err = 'sorry, pdf already exists';
				res.status(400).json(err);
				return;
			}
			
			//update database entry
			Presentation.find({ 'pres_ID' : req.params.pres_ID })
			.exec(function(err, pres) {
				if (err) {
					res.status(400).json(err);
					return;
				}
				if (!pres[0]) {
					res.status(400).json('did not find presention');
					return;
				}
				var r_pres = pres[0];
				r_pres.download = true;
				var now = new Date();
				var updated = now.toJSON();
				r_pres.updated = updated;
				r_pres.save(function(err) {
					if (err) {
						res.status(400).json(err);
						return;
					}else{
						fs.rename('./uploads/' + old_filename, './public/files/' + pres_ID + '/presentation.pdf');
						res.status(201).json('upload succeeded!');
					}
				});
		
			});
		}
	});
});

// call for checking if successful connection can be established
app.get('/api/v1/presentations/verify', function(req, res) {
		res.status(200);
});

// presentation tracking
app.get('/track/:pres_ID', function(req, res) {
	res.sendfile('./public/track.html');
});

// iframe tracking
app.get('/embed/:pres_ID', function(req, res) {
  res.sendfile('./public/embed.html');
});

// download presentation tool
app.get('/download', function(req, res) {
	res.sendfile('./public/download.html');
});

app.get('/download/slideTracker', function(req, res) {
	var file = './public/files/slideTracker_0_1_1_0.zip';
	res.download(file);
});

// api-docuentation
app.get('/api-documentation', function(req, res) {
	res.sendfile('./public/api.html');
});

// terms & privacy policy
app.get('/privacy', function(req, res) {
	res.sendfile('./public/privacy.html');
});

// contact
app.get('/contact', function(req, res) {
	res.sendfile('./public/contact.html');
});

// admin section callbacks
app.get('/admin/get_log_entries', function(req, res) {
	LogPres.find().exec(function(err, presentations) {
		if (err) {
			return next(err)
		}
		res.status(200).json(presentations)
	})
});

// admin section callbacks
app.get('/admin/get_db_entries', function(req, res) {
	Presentation.find().exec(function(err, presentations) {
		if (err) {
			return next(err)
		}
		res.status(200).json(presentations)
	})
});

app.get('/admin/analytics', function(req, res) {
	res.sendfile('./admin/analytics.html');
});

app.get('/admin/overview', function(req, res) {
	res.sendfile('./admin/admin.html');	
});

// export log files in CSV format
app.get('/admin/export', function(req, res) {
	LogPres.find().exec(function(err, presentations) {
		if (err) {
			return next(err)
		}
		json2csv(
	    {   data: presentations, 
	        fields: ['unique_ID','clients','active','download','n_slides','file_size','created'], 
	        fieldNames: ['unique_ID','clients','active','download','n_slides','file_size','created']
	    }, 
	    function(err, csv) {  
	        if (err) console.log(err);
			res.setHeader('Content-disposition', 'attachment; filename=slideTracker_logs.csv');
			res.setHeader('Content-type', 'text/csv');
			res.send(csv);
	    });   
	})
});


// app.get('/test/s3', function(req, res) {
// // var s3 = new AWS.S3();
// // var params = {Bucket: 'slide-tracker-s3', Key: 'myImageFile.jpg'};
// // var file = fs.createWriteStream('./uploads/file.jpg');
// // s3.getObject(params).createReadStream().pipe(file);
// var s3 = new AWS.S3({params: {Bucket: 'slide-tracker-s3', Key: 'myKey'}});
// s3.createBucket(function() {
  // s3.upload({Body: 'Hello!'}, function() {
    // console.log("Successfully uploaded data to myBucket/myKey");
  // });
// });
// res.status(200);
// });

// robots
app.get('/robots.txt', function(req, res) {
	res.sendfile('./public/robots.txt');
});

app.post('/status', function(req, res) {
	var response = {
			error : 0,
			message : 'you have the newest version running'
		};
	res.status(200).json(response);
});

app.get('*', function(req, res) {
	res.sendfile('./public/index.html');
});

io.on('connection', function(socket) {
	// user connected
	if (socket.handshake.query.pres_ID){
		Presentation.find({ 'pres_ID' : socket.handshake.query.pres_ID })
		.exec(function(err, pres) {
			if (err) {return;}
			if (!pres[0]) {return;}
			var r_pres = pres[0];					
			r_pres.clients = r_pres.clients + 1;
			
			// update log entry if higher peak users
			LogPres.find({ 'unique_ID' : r_pres.pres_ID+'_'+r_pres.created.toISOString()})
			.exec(function(err, lpres) {
				var log_pres = lpres[0];
				if(r_pres.clients > log_pres.clients){
					log_pres.clients = r_pres.clients;
					log_pres.save(function(err) {})
				}
			});
			
			r_pres.save(function(err) {if (err) {return;}});
		});
	}
socket.on('disconnect', function() {
	// user disconnected
	if (socket.handshake.query.pres_ID){
		Presentation.find({ 'pres_ID' : socket.handshake.query.pres_ID })
		.exec(function(err, pres) {
			if (err) {return;}
			if (!pres[0]) {return;}
			var r_pres = pres[0];					
			r_pres.clients = r_pres.clients - 1;
			r_pres.save(function(err) {if (err) {return;}});
		});
	}
});
});


https.listen(app.get('port'), function() {
	console.log('Express server listening on port ' + app.get('port'));
});

// http redirect
var app2 = express();
var http = require('http').Server(app2);

// set up a route to redirect http to https
app2.get('*',function(req,res){  
	//res.redirect('https://www.slidetracker.org'+req.url)
    //res.redirect('https://dev.slidetracker.org'+req.url)
    res.redirect('https://localhost:3000'+req.url)
})

// have it listen on 8080
http.listen(8080);

