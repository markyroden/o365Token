
(function () {
    'use strict';
    var token = app.getCookie("OAuthToken")
    //check to see if there is an OAuth Token cached in the cookie
    //we check the URL because the first time the page is returned from Authorization
    //the cookie has not been set - without this check we would open extra windows......
    if (!token && (location.href.split('access_token').length<2)){
        $(document).ready(function () {
            if (!token) {
                var tokenParams = {};
                tokenParams.authServer = 'https://login.windows.net/common/oauth2/authorize?';
                tokenParams.responseType = 'token';
                //where the OAuth process sends the web page back to once the token
                //has been successfully generated
                tokenParams.replyUrl = "https://xomino365.com";

                //THESE tokenParams need to be changed for your application
                tokenParams.clientId = '1bc6eb91-a9b4-47ee-9add-7d14ff5d77bd';
                //this is the O365 site from where you are going to get data
                tokenParams.resource = "https://xomino.sharepoint.com";

                app.getToken(tokenParams);
            } else {
                //we have a token therefore carry on
                app.tokenCallback(token)
            }
        });
    } else {
        //The window has opened for token generation
        $(document).ready(function () {
            app.returnToken();
        });
    }
})();


