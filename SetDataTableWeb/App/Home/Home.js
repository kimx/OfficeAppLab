/// <reference path="../App.js" />

(function () {
    "use strict";

    // The initialize function must be run each time a new page is loaded
    Office.initialize = function (reason) {
        $(document).ready(function () {
            app.initialize();

            $('#set-data-from-selection').click(setDataTable);
        });
    };

    // Reads data from current document selection and displays a notification
    function setDataTable() {
        var dataTable = new Office.TableData();
        dataTable.headers = ['Name', 'City'];
        dataTable.rows = [['Kim', 'Kao'], ['David', 'Tai'], ['John', 'Tan']];
        Office.context.document.setSelectedDataAsync(dataTable, { CoercionType: Office.CoercionType.Table },
            function (result) {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    app.showNotification('The selected text is:', '"' + result.status + '"');
                } else {
                    app.showNotification('Error:', result.error.message);
                }
            }
        );
    }
})();