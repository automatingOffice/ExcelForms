(function () {
    "use strict";

    // The initialize function must be run each time a new page is loaded.
    Office.initialize = function (reason) {
        $(document).ready(function () {
            $('#submit-form').click(submitData);
        });
    };

    function submitData() {
        Excel.run(function (ctx) {
            //password is testing
            const worksheetProtection = ctx.workbook.worksheets.getItem('clientList').protection; 
            worksheetProtection.unprotect('testing');
            var dataRange = ctx.workbook.worksheets.getItem('clientList').getRange('A1').getSurroundingRegion();
            dataRange = dataRange.getResizedRange(1);
            const lastRow = dataRange.getLastRow(); 
            lastRow.values = [[$('#input-name').val(), $('#input-email').val(), $('#input-department').val()]]
            return ctx.sync().then(function () {
                worksheetProtection.protect(null,'testing');
                return ctx.sync();
            });
        })
            .catch(function (error) {
                console.log("Error: " + error);
                if (error instanceof OfficeExtension.Error) {
                    console.log("Debug info: " + JSON.stringify(error.debugInfo));
                }
            });
    }



})();