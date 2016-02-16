
(function () {
    'use strict';

    // The Office initialize function must be run each time a new page is loaded
    if (location.href.split("Outlook").length > 1){
        Office.initialize = function (reason) {
            $(document).ready(function () {
                app.initialize();
            });
        };
    } else {
        //testing the app not inside of Outlook
        $(document).ready(function () {
            app.initialize();
        });
    }


})();