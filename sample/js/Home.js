
(function () {
    'use strict';

    // The Office initialize function must be run each time a new page is loaded in the Add-In
    if (location.href.split("host_Info").length > 1){
        Office.initialize = function (reason) {
            $(document).ready(function () {

                var tokenParams = {}
                tokenParams.authServer   = 'https://login.windows.net/common/oauth2/authorize?';
                tokenParams.responseType = 'token';

                //THESE tokenParams need to be changed for your application
                tokenParams.clientId     = 'Your-app-clientId-goes-here'
                tokenParams.replyUrl     = location.href.split("?")[0];
                tokenParams.resource	 = "https://yoursite.sharepoint.com"

                app.getToken(tokenParams);

            });
        };
    } else {
        //The window has opened for token generation
        $(document).ready(function () {

            app.returnToken();

        });
    }


})();