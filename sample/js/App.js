// Common app functionality

var app = (function () {
    'use strict';

    var app = {};

    // Function called when the page has opened outside of the Add-In
    app.getToken = function (tokenParams) {

        var url =   tokenParams.authServer +
                    "response_type=" + encodeURI(tokenParams.responseType) + "&" +
                    "client_id=" + encodeURI(tokenParams.clientId) + "&" +
                    "resource=" + encodeURI(tokenParams.resource) + "&" +
                    "redirect_uri=" + encodeURI(tokenParams.replyUrl);

        var winObj = window.open(url);
        winObj.focus();

    };

    app.returnToken = function(){
        var urlParameterExtraction = new (function () {
            function splitQueryString(queryStringFormattedString) {
                var split = queryStringFormattedString.split('&');

                // If there are no parameters in URL, do nothing.
                if (split == "") {
                    return {};
                }

                var results = {};

                // If there are parameters in URL, extract key/value pairs.
                for (var i = 0; i < split.length; ++i) {
                    var p = split[i].split('=', 2);
                    if (p.length == 1)
                        results[p[0]] = "";
                    else
                        results[p[0]] = decodeURIComponent(p[1].replace(/\+/g, " "));
                }
                return results;
            }

            // Split the query string (after removing preceding '#').
            this.queryStringParameters = splitQueryString(window.location.hash.substr(1));
        })();

        var token = urlParameterExtraction.queryStringParameters['access_token'];

        if (window.opener){
            window.opener.app.tokenCallback(token);
            window.close()
        } else {
            document.write("There appears to have been an error, please close the window and check with your administrator")
        }

    }

	app.tokenCallback = function(token){
		//this would then continue to do what you really need in the Add-In
        document.getElementById('tokenHere').innerHTML = token
	}


    return app;
})();