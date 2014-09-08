ParameterizationPreview
=======================

Overview
--------

This Visual Studio extension provides the easiest way to preview MSDeploy (WebDeploy) parameterization against an applications config file.  The extension currently supports VS2012 and VS2013.

To use the tool simply right click on a "SetParamaters.[configuration].xml" (i.e. SetParameters.Release.xml) file in the Solution Explorer and choose the "ParameterizationPreview" option.  A few command windows will pop up then close and eventually display the comparison of the original config file against the parameterized config.

Behind the scenes a MSDeploy package is created from the config files and then deploys that package to a temp directory at the root of the solution.  Finally a VS file comparison of the resulting parameterized config against the original is displayed.

Troubleshooting
---------------

In some situations there may be an issue with the creation of the temporary MSDeploy package or parameterizing the config file.  When this occurs the preview will fail with few details.  You can run the preview in troubleshooting mode by holding down the control (CTRL) button when right clicking the ParameterizationPreview option.  This will leave each command window open to review its contents for errors/information and then you must manually close each window.

MSDeploy Parameterization Links
-------------------------------

Here are a few good posts about Parameterization:

- http://www.iis.net/learn/publish/using-web-deploy/web-deploy-parameterization
- http://blogs.iis.net/elliotth/archive/2012/12/17/web-deploy-xml-file-parameterization.aspx
