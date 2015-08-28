/// <reference path="../App.js" />

(function () {
    "use strict";
    // The initialize function must be run each time a new page is loaded
    Office.initialize = function (reason) {
        $(document).ready(function () {
            app.initialize();
            $('#get-data-from-selection').click(getDataFromSelection);
            $('#set-data-from-selection').click(setDataFromSelection);
        });
    };

    function getDataFromSelection() {
        //取得選取的儲存格資料，Office.CoercionType.Text 指定回傳結果的格式(本例為文字)
        Office.context.document.getSelectedDataAsync(Office.CoercionType.Text,
            function (result) {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    app.showNotification('選取的資料:', '"' + result.value + '"');
                } else {
                    app.showNotification('錯誤:', result.error.message);
                }
            }
        );
    }

    function setDataFromSelection() {
        Office.context.document.setSelectedDataAsync($("#myValue").val(),null,
            function (result) {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    app.showNotification('設定結果:', '"' + result.status + '"');
                } else {
                    app.showNotification('錯誤:', result.error.message);
                }
            }
        );
    }
})();