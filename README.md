# slideTracker
Welcome to our slideTracker dev site. Since we just just started developing we appreciate help and discourage harsh criticism :) We are operating under MIT license, so please feel free to fork and contribute as much as you want!

## Our Production Server:
[http://www.slidetracker.org](http://www.slidetracker.org)

## Our Development Server:
[http://54.208.192.158](http://54.208.192.158)

## Getting started developing
SlideTracker consists of two parts; a PowerPoint add-in acting as a client and a <a>MEAN stack</a> driven web-app to provide a server with a RESTful [API](http://54.208.192.158/api-documentation) and a simple web interface for users. The source code for both can be found in [our repository](https://github.com/GeorgKucsko/slideTracker). Depending on what you wish to be working on you need a different setup. Also check out our [dev server](http://54.208.192.158) where we will try to always have the newest code commit running for you to test out. 

## Developing the slideTracker PowerPoint add-in
The add-in was developed using Visual Studio 2013 in C#. Some version of Visual Studio are available for free online for individuals (read more [here](http://www.visualstudio.com/products/visual-studio-community-vs)). Simply copy the files from the add-in folder of our repository and import the project into your Visual Studio workspace. We try to provide good inline documentation within our code. To find out what endpoints our API offers check [the documentation](http://54.208.192.158/api-documentation). Please use our github issue-tracker to see active areas of development or to contact us with ideas. 

**NOTE:** In ThisAddIn.cs, there is a property postURL. There are two versions of the URL- one for the dev server and one for the production server. Please develop and debug using the dev server and leave the production server commented out. 

## Developing the slideTracker web-app
If you want to play around with the server code, make sure you have node.js installed on your system. Next copy the repository node.js folder to your drive and open up the app.js file in your favorite editor. To store presentation date we use a [mongolab](https://mongolab.com) MongoDB database. Simply exchange the corresponding line in the code with your database credentials (we use the environment variable “MONGODB_STRING”). The main angular functions for presentation viewing can be found in public/core.js. We try to provide good inline comments on our code, but if something is unclear please post an issue on github and we will try to improve clarity.

Have fun developing!




