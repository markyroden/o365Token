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

* clientid
* resource

For more information check out the blog post - http://xomino365.com/2016/02/28/o365token-github-project-authenticating-office-add-ins-to-azure-ad-using-only-javascript-no-c/
