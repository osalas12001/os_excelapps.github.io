(function () {
    "use strict";

    var cellToHighlight;
    var messageBanner;
    var htmlWrite = "";
    





    // The initialize function must be run each time a new page is loaded.
    Office.initialize = function (reason) {
        $(document).ready(function () {
            // Initialize the FabricUI notification mechanism and hide it
            var element = document.querySelector('.ms-MessageBanner');
            
            messageBanner = new fabric.MessageBanner(element);
            messageBanner.hideBanner();
            
            // If not using Excel 2016, use fallback logic.
            if (!Office.context.requirements.isSetSupported('ExcelApi', '1.1')) {
               $("#template-description").text("This sample will display the value of the cells that you have selected in the spreadsheet.");
               $('#button-text').text("Submit!");
               $('#button-desc').text("Submit project Information.");

                $('#highlight-button').click(writeData);
                return;
            }

            $("#template-description").text("This sample highlights the highest value from the cells you have selected in the spreadsheet.");
            $('#button-text').text("Submit");
            $('#button-desc').text("Submit project information.");
                

            loadAutocomplete();
            populateProjStatus();
            initDatePicker();
            

            // Add a click event handler for the highlight button.
            $('#highlight-button').click(writeData);
        });
    };

    
   
    // return integer of empty row;  this  will define what row to write data to 
    function writeData() {
        //define the autoomplete array to load
        var values = [];
        var i = 0;
        var lastRow = 0;
        // read from the desired workbook
       
        Excel.run(function (ctx) {

            
            var sheet = ctx.workbook.worksheets.getActiveWorksheet();
            var range = sheet.getRange("D6:L1000").load("values, rowCount");
            return ctx.sync().then(function () {
                // load the list values from the spreadsheet
                for (i = 0; i < range.rowCount; i++) {
                    values[i] = range.values[i].toString();
                   // OS-Modify commas needed for this data when blank or figure out how to strip commas; 
                   // or check if the first column if each row blank.  Maybe this is best
                  // check if all form elements have been defined else, notify user through shownotification that 
                  // missin data must be defined before we can submit

                    console.log(values[i]);
                    if ((values[i] === ",,,,,,,,") || (values[i] === ",,,,,,0,,") ) {
                        lastRow = i;
                        
                        dataToWrite(lastRow);
                        break;
                    }
                }
            }).then(ctx.sync);
        }).catch(errorHandler);
    }

    function dataToWrite(rowNumber)
    {
        var clientName, projectName, projectAddress, projectNumber, projectStatus, projectFee, bonusPercentage, projectYear, timestamp; 
        var row, offsetRow; 
        var inputData = [];

        clientName = $('#client-name').val();
        projectName = $('#project-name').val();
        projectAddress = $('#project-address').val();
        projectNumber = $('#project-number').val();
        projectStatus = $('#project-status').val();
        projectFee = $('#project-fee').val();
        bonusPercentage = $('#project-bonus').val();
        projectYear = $('#project-year').val();
        timestamp = $('#time-stamp').val();

        

        offsetRow = 5;
        row = rowNumber + offsetRow;
        var dataRange = "D6:L1000";

        // before we continue check to see if any fields are empty;  if so nojtify user to 
        // fill out all information
        inputData = [clientName, projectName, projectAddress, projectNumber, projectStatus, projectFee, bonusPercentage, projectYear, timestamp];

        for (var i = 0; i < inputData.length; i++){

            if (emptyInputCheck(inputData[i])) {
                showNotification("Please fill out all fields"); 
               
                return;
            }
          
        }



        // before we write data; make sure all fields are not blank. 
        // else show a notification to the user through the showNotification function

        Excel.run(function (ctx) {
            var sheet = ctx.workbook.worksheets.getActiveWorksheet();
            var range = sheet.getRange(dataRange);
            // os create a switch statement for each column such that text field updates the 
            // correct column value
            return ctx.sync().then(function () {

                sheet.getCell(row, 3).values = clientName;
                sheet.getCell(row, 4).values = projectName;
                sheet.getCell(row, 5).values = projectAddress;
                sheet.getCell(row, 6).values = projectNumber;
                sheet.getCell(row, 7).values = projectStatus;
                sheet.getCell(row, 8).values = projectFee;
                sheet.getCell(row, 10).values = bonusPercentage;
                sheet.getCell(row, 11).values = projectYear;
                sheet.getCell(row, 13).values = timestamp;
            }).then(ctx.sync);

        }).catch(errorHandler); 
    }

    function populateProjStatus()
    {
        var status = [];
        var address = "Y7:Y11";
        var htmlWrite = "";
        var select = document.getElementById("project-status");
        
        Excel.run(function (ctx) {
            var sheet = ctx.workbook.worksheets.getActiveWorksheet();
            var range = sheet.getRange(address).load("values, rowCount");
            return ctx.sync().then(function () {
                status = range.values;
                for (var i = 0; i < range.rowCount; i++){
                    htmlWrite += "<option>" + status[i] + "</option>"; 
                }
                select.innerHTML = htmlWrite;
                initJobStatus();
              
            }).then(ctx.sync);
        }).catch(errorHandler);
    }

    function initJobStatus() {

        if ($.fn.Dropdown) {
            $('.ms-Dropdown').Dropdown();
        }
    }
    function initDatePicker()
    {  
        if ($.fn.DatePicker) {
            $('.ms-DatePicker').DatePicker();
        }   
    }

    // check that non of the input boxes are empty; else return bool true=not-all are filled; false = all-filled
    function emptyInputCheck(input)
    {
        if (input === ""){
            return true;
        }
        else 
        {
            return false;
        }
    }

    function loadAutocomplete()
    {
        // Load the list of clients currently saved 
        var clientArray = []; 


        //Lets load data from the workbook
        Excel.run(function (ctx) {
            var sheet = ctx.workbook.worksheets.getActiveWorksheet();
            var range = sheet.getRange("Z7:Z100").load("values, rowCount");

            return ctx.sync().then(function () {
                // load the list values from the spreadsheet
                for (var i = 0; i < range.rowCount; i++) {
                    //console.log(range.values[i].toString());
                    clientArray[i] = range.values[i].toString();
                    if (clientArray[i] === "") { break; }

                }
                
                $('#client-name').autocomplete({
                    source: clientArray
                });
            }).then(ctx.sync);

        }).catch(errorHandler);
    }

    // Helper function for treating errors
    function errorHandler(error) {
        // Always be sure to catch any accumulated errors that bubble up from the Excel.run execution
        showNotification("Error", error);
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
    }

    // Helper function for displaying notifications
    function showNotification(header, content) {
        $("#notification-header").text(header);
        $("#notification-body").text(content);
        messageBanner.showBanner();
        messageBanner.toggleExpansion();
    }
})();
