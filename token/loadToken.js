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
