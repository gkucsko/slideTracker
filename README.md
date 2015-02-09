# slideTracker
Welcome to our slideTracker dev site. Since we just just started developing we appreciate help and discourage harsh criticism :) We are operating under the MIT license, so please feel free to fork and contribute as much as you want!

## What is SlideTracker
SlideTracker provides a convenient way for a PowerPoint presentation to be broadcasted online. This way the audience can follow along on their own device and navigate through the slides independently if needed. All this is happening with as little set-up as possible. All it takes on the presenter side is to install our PowerPoint add-in and a single click to start broadcasting. For the audience it simply requires a browser (mobile or otherwise) and a unique presentation ID displayed on the presenter's slides.
Below please find a quick summary on the architecture we are using and some pointers on how to get you started on contributing. Additionally our code should be commented relatively well inline. 

## Our Production Server (for anyone actually using slideTracker):
[http://www.slidetracker.org](http://www.slidetracker.org)

## Our Development Server (always updated with current development version):
[http://dev.slidetracker.org](http://dev.slidetracker.org)

## Getting started developing
SlideTracker consists of two parts; a PowerPoint add-in acting as a client and a [MEAN stack](http://en.wikipedia.org/wiki/MEAN) driven web-app to provide a server with a RESTful [API](http://dev.slidetracker.org/api-documentation) and a simple web interface for users. The source code for both can be found in [our repository](https://github.com/GeorgKucsko/slideTracker). Depending on what you wish to be working on you need a different setup. 

## Developing the slideTracker PowerPoint add-in
The add-in was developed using Visual Studio 2013 in C#. Some version of Visual Studio are available for free online for individuals (read more [here](http://www.visualstudio.com/products/visual-studio-community-vs)). Simply copy the files from the add-in folder of our repository and import the project into your Visual Studio workspace. To find out what endpoints the slideTracker web API offers check [the documentation](http://dev.slidetracker.org/api-documentation). Please use our github issue-tracker to see active areas of development or to contact us with ideas. 

**NOTE:** In ThisAddIn.cs, there is a property postURL. Make sure to substitute the correct URL (localhost or dev server). Please develop and debug using the dev or a localhost server and leave the production server commented out. 

## Developing the slideTracker web-app
If you want to play around with the server code, make sure you have node.js installed on your system. Next, copy the repository node.js folder to your drive and open up app.js (main server file) in your favorite editor. As a database we use a [mongolab](https://mongolab.com) MongoDB database. Simply exchange the corresponding line in the code with your database credentials (we use the environment variable “MONGODB_STRING”). Additionally there is an admin section for which you can define a password with the environment variable “ADMIN_PW“, however for development purposes this section can be ignored.
The main angular functions for presentation viewing can be found in public/core.js. We try to provide good inline comments on our code, but if something is unclear please post an issue on github and we will try to improve clarity.

Have fun developing!




