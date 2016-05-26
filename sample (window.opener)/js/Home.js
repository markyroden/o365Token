
(function () {
    'use strict';
    // The Office initialize function must be run each time a new page is loaded in the Add-In
    if (location.href.split("host_Info").length > 1){
        Office.initialize = function (reason) {
            $(document).ready(function () {
                //check to see if there is an OAuth Token cached in the cookie
                //app.addinName will need to be different for every Add-In
                var token = app.getCookie(app.addinName)
                if (!token) {
                    var tokenParams = {};
                    tokenParams.authServer = 'https://login.windows.net/common/oauth2/authorize?';
                    tokenParams.responseType = 'token';
                    tokenParams.replyUrl = location.href.split("?")[0];

                    //THESE tokenParams need to be changed for your application
                    tokenParams.clientId = 'Your-app-clientId-goes-here';
                    tokenParams.resource = "https://yoursite.sharepoint.com";

                    app.getToken(tokenParams);
                } else {
                    //we have a token therefore carry on
                    app.tokenCallback(token)
                }
            });
        };
    } else {
        //The window has opened for token generation
        $(document).ready(function () {
            app.returnToken();
        });
    }
})();


