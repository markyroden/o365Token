# o365Token
Sample application code for generating an Authorization token from O365.
The project initially focused on Office 365 Add-ins but was extended to cover O365 Authentication and Authorization from anywhere.

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

For more information check out the blog post - https://xomino.com/

# Release notes

Version 1.0
* This release demonstrates the ability to create an Authorization key within the Outlook Add-in itself. There is no need for using a Dialog or window.opener.


Version 0.3
* Added O365Chrome_Extension Example

Version 0.2
* Added document.cookie caching.
* New parameter also created for passing Add-In name to the Cookie getters and setters
* The token is now cached for 3600 seconds as a cookie in the domain of the hosted add-in. That means if the user needs to use the Add-In again within 60 minutes the popup window will not appear.
* This cookie needs to be set for each Add-In because the necessary permissions and end-points will always be different

Version 0.1
* Initial release
For more information check out the blog post - http://xomino365.com/2016/02/28/o365token-github-project-authenticating-office-add-ins-to-azure-ad-using-only-javascript-no-c/


# License

The code in this project governs my unique contributions and is open, free to use and redistribute or modify, for public or private use, without warranty or liability.

The MIT License (MIT) © 2016 Mark Roden

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.