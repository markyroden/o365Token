# o365Token
Sample application code for generating an Authorization token from within O365 Office Add-ins

# What it does
This sample application will allow you ro return an OAuth Authorization token for use with O365. What you do with it will depend on the set up.
The project assumes that you have your O365 environment associated with Azure AD as your directory source.

# Instructions
Clone the repository
Create your Azure AD site and make the redirect URL match the sample/index.html path.

e.g. https://yoursite.azurewebsites.net/sample/index.html

Modify the following token parameters in app.js to match your application

  clientid
  
  resource
