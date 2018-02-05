# createuploadsession-graph-jquery
Code example to show how to upload larger files to either SharePoint Online or OneDrive for Business document libraries using plain JavaScript and jQuery.

To run the sample:
 - Create an Azure AD Application at https://apps.dev.microsoft.com
 - Configure the a Redirect URL to match your development environmnet (e.g. http://localhost:5500.  Note, some browsers may send a trailing '/' so best to add the form as well).
 - Enable 'Allow Implicit Flow.'
 - Give Files.ReadWrite permission scope.
 - Change values of the "clientId" and the "tenant" variable in the config.js file.
