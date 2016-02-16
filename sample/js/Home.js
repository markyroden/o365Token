
(function () {
    'use strict';

    // The Office initialize function must be run each time a new page is loaded
    if (location.href.split("host_Info=Outlook").length > 1){
        Office.initialize = function (reason) {
            $(document).ready(function () {

                var tokenParams = {}
                // Creat the token parameters
                tokenParams.authServer  = 'https://login.windows.net/common/oauth2/authorize?';
                tokenParams.clientId    = document.getElementById('clientId').value;
                tokenParams.replyUrl    = location.href.split("?")[0]; //document.getElementById('replyURL');
                tokenParams.resource	= document.getElementById('resource').value;
                tokenParams.responseType = 'token';

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