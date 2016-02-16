// Common app functionality

var app = (function () {
    'use strict';

    var app = {};

	app.tokenCallback = function(token){
		//this would then continue to do what you really need in the Add-In
        document.getElementById('tokenHere').innerHTML = token
	}

    // Common initialization function (to be called from each page)
    app.initialize = function () {
       
        // found on the Configure tab in the Azure Management Portal.
        var clientId    = document.getElementById('clientId');
        var replyUrl    = document.getElementById('replyURL');
        var resource	= document.getElementById('resource');
        var authServer  = 'https://login.windows.net/common/oauth2/authorize?';

        var responseType = 'token';

        var url = authServer +
            "response_type=" + encodeURI(responseType) + "&" +
            "client_id=" + encodeURI(clientId) + "&" +
            "resource=" + encodeURI(resource) + "&" +
            "redirect_uri=" + encodeURI(replyUrl);

           var winObj = window.open(url);
           winObj.focus();
	   
    };

    return app;
})();