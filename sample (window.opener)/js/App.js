// Common app functionality

var app = (function () {
    'use strict';

    var app = {};

    //Change this value for every Add-In
    app.addinName = "OAuthToken"

    // Function called when the page has opened in the Outlook Add-In
    app.getToken = function (tokenParams) {

        var url =   tokenParams.authServer +
            "response_type=" + encodeURI(tokenParams.responseType) + "&" +
            "client_id=" + encodeURI(tokenParams.clientId) + "&" +
            "resource=" + encodeURI(tokenParams.resource) + "&" +
            "redirect_uri=" + encodeURI(tokenParams.replyUrl);

        var winObj = window.open(url);
        winObj.focus();

    };

    //Function called when the page has been called post token-generation in the pop-up window
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
            app.setCookie(token, 3600);
            window.opener.app.tokenCallback(token);
            window.close()
        } else {
            document.write("There appears to have been an error, please close the window and check with your administrator")
        }

    };

    //The code which runs the Add-In itself
    app.tokenCallback = function(token){
        //this would then continue to do what you really need in the Add-In
        document.getElementById('tokenHere').innerHTML = token
    };

    app.setCookie = function(value, duration) {
        var now = new Date();
        var time = now.getTime();
        var expireTime = time + 1000*duration;
        now.setTime(expireTime);
        console.log("---setting cookie--");
        document.cookie = app.addinName+"="+value+';expires='+now.toGMTString()+';path=/';
        console.log(document.cookie);
    }

    app.getCookie = function (addinName, name){
        var pattern = RegExp(name + "=.[^;]*")
        var matched = document.cookie.match(pattern)
        if(matched){
            var cookie = matched[0].split('=')
            return cookie[1]
        }
        return false
    }


    return app;
})();

