var months = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"]

function doGet() {
  var currentYear = new Date().getFullYear(); // Get the current year
  var html = HtmlService.createHtmlOutput(`
    <!DOCTYPE html>
    <html>
      <head>
      <style>
          body {
            font-family: 'Arial', sans-serif; /* Modern non-serif font */
            line-height: 1.6;
            background-color: #f0f4f8;
            margin: 0;
            padding: 0;
          }
          h1 {
            text-align: center;
            color: #333;
          }
          form {
            max-width: 400px;
            margin: 0 auto;
            padding: 20px;
            background-color: #ffffff;
            border-radius: 8px; /* Rounded corners */
            box-shadow: 0 4px 8px rgba(0,0,0,0.1);
          }
          label {
            display: block;
            margin: 10px 0 5px;
            font-weight: bold;
          }
          input, select, button {
            width: 100%;
            padding: 10px;
            margin: 5px 0 10px;
            border: 1px solid #ddd;
            border-radius: 8px; /* Rounded corners */
            box-sizing: border-box;
          }
          button {
            background-color: #007BFF;
            color: white;
            font-size: 16px;
            cursor: pointer;
          }
          button:hover {
            background-color: #0056b3;
          }
        </style>
        <title>Tracker Calendar Creator</title>
        <script>
          let year, sheetId, goal, sheetName;

          function handleFormSubmit() {
            // Get values from input fields
            year = document.getElementById("year").value;
            sheetId = document.getElementById("sheetId").value;
            goal = document.getElementById("goal").value;
            sheetName = document.getElementById("sheetName").value || "Sheet1";

            // Call the Apps Script function and pass the values
            google.script.run
              .withSuccessHandler(() => {
                console.log("Calendar creation started successfully with Year:", year, "Sheet ID:", sheetId, "Sheet Name:", sheetName);
                alert("Calendar creation started!");
              })
              .withFailureHandler((error) => {
                console.error("Error occurred: ", error);
                alert("Error: " + error.message);
              })
              .workoutCalendarCreator(sheetId, year, goal, sheetName);
          }
        </script>
      </head>
      <body>
        <h1>Tracker Calendar</h1>
        <form id="inputForm" onsubmit="event.preventDefault(); handleFormSubmit();">
          <label for="sheetId">Google Sheet ID:</label><br>
          <input type="text" id="sheetId" name="sheetId" required><br><br>
          <label for="sheetName">Sheet Name:</label><br>
          <input type="text" id="sheetName" name="sheetName" placeholder="Sheet1"><br><br>
          <label for="year">Year:</label><br>
          <input type="number" id="year" name="year" min="1900" max="3000" value="${currentYear}" required>
          <label for="goal">Weekly goal:</label><br>
          <select id="goal" name="goal" required>
            <option value="1">1</option>
            <option value="2">2</option>
            <option value="3">3</option>
            <option value="4" selected>4</option>
            <option value="5">5</option>
            <option value="6">6</option>
            <option value="7">7</option>
          </select><br><br>
          <button type="submit">Create Calendar</button>
        </form><br>
        <form> Hello user! &#128522 This is simple (but effective!) weekly habit tracker. This project started as a personal solution to help me track my workout routine habit. I didn't want to rely on a third party app, so I created a simple spreadsheet to keep everything organized, and after encountering a few bugs and sharing the solution on Reddit, I was amazed by the positive response and the requests for access to the file. This encouraged me to turn my custom code into something more general, allowing others to adapt it to suit their own needs. After discovering Google Apps Script’s ability to create a web app, I decided to share it in this format, making it easy for users to generate a template in their Google Sheets to track their progress.

      I’d love to hear any feedback, bugs, or ideas for improvements! I hope this helps you stay on top of your weekly goals and track your progress!  </form>
      </body>
    </html>
  `);
  return html;
}

function workoutCalendarCreator(sheetId, year=2025, goal=4, sheetName="Sheet1") {
  var spreadsheet = SpreadsheetApp.openById(sheetId)
  var sheet = spreadsheet.getSheetByName(sheetName)
  resetSheet(sheet);

  var januaryOriginRow = 9
  var januaryOriginCol = 2

  var averageTimesPerWeekRange = []
  var weekNumberRange = []
  var prevMonth = null
  var lastWeekOfTheYear
  for(var i = 1; i <= 12; i++){
    var monthNumber = i
    var monthStruct = getMonthStruct(monthNumber, year)

    var originRow = januaryOriginRow+8*Math.floor((monthNumber-1)/2)
    var originCol = (monthNumber % 2 === 1) ? januaryOriginCol : januaryOriginCol + 16;
       
    createGrid(sheet, originRow, originCol, monthNumber, year)
    const result = addDaysPerWeekCount(sheet, originRow, originCol, monthStruct, averageTimesPerWeekRange, prevMonth)
    prevMonth = result.sumRange
    averageTimesPerWeekRange = result.averageTimesPerWeekRange
    weekNumberRange, lastWeekOfTheYear = addWeekNumber(sheet, originRow, originCol, monthStruct, monthNumber, weekNumberRange, year)
    
  
    // Set conditional format rules
    var rules = sheet.getConditionalFormatRules();
    var cellRange = getRangeForMonth(originRow, originCol, monthNumber, year)
    for (var j = 0; j < cellRange.length; j++){
      var cellWithNumber = sheet.getRange(cellRange[j][0],cellRange[j][1])
      var cellWithCheckbox = sheet.getRange(cellRange[j][0],cellRange[j][1]+1)

      var cellString = convertToCellReference(cellRange[j][0],cellRange[j][1]+1)
      var ruleCheckbox = SpreadsheetApp.newConditionalFormatRule()
                                .whenFormulaSatisfied('=$'+cellString)
                                .setBackground('#b7e1cd')
                                .setRanges([cellWithNumber, cellWithCheckbox])
                                .build();
      var cellDate = "=DATE("+year+","+String(i)+","+String(cellWithNumber.getValue())+")"
      var ruleBackground = SpreadsheetApp.newConditionalFormatRule()
                                        .whenFormulaSatisfied(cellDate+" <= TODAY()")
                                        .setBackground("#ff9797")
                                        .setRanges([cellWithNumber, cellWithCheckbox])
                                        .build();
      rules.push(ruleCheckbox)
      rules.push(ruleBackground)
    }
    sheet.setConditionalFormatRules(rules)
  }

  //const goal = "208"
  //const goal1 = timesPerWeek * lastWeekOfTheYear;
  addHeaderFormulas(sheet, goal, averageTimesPerWeekRange, lastWeekOfTheYear);

}

function resetSheet(sheet) {
  try {
    // Get the range of the sheet
    var range = sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns());

    // Clear all cell values, formats, and contents
    range.clearContent();
    range.clearFormat();

    // Remove conditional formatting rules
    sheet.clearConditionalFormatRules();

    // Remove all borders
    range.setBorder(false, false, false, false, false, false);

    // Remove all data validations (including checkboxes)
    range.clearDataValidations();

    // Resize all columns to the default width (100)
    for (var i = 1; i <= sheet.getMaxColumns(); i++) {
      sheet.setColumnWidth(i, 100);
    }

    // Set the background color to white (or transparent)
    range.setBackground(null); // null resets to the default color, which is white/transparent

    // Remove pinned (frozen) rows
    sheet.setFrozenRows(0);

    // Ensure there are at least 33 columns
    var currentColumns = sheet.getMaxColumns();
    if (currentColumns < 33) {
      sheet.insertColumnsAfter(currentColumns, 33 - currentColumns);
    }

  } catch (error) {
    Logger.log(`Error resetting sheet: ${error.message}`);
    throw error;
  }
}

function getMonthStruct(monthNumber, year) {
  //var realFirstDay = new Date(year, monthNumber-1, 7).getDay()
  var firstDay = new Date(year, monthNumber-1, 1).getDay()
  //Readjust so that monday is 0 and sunday is 6
  firstDay = (firstDay + 6) % 7
  var daysInMonth = new Date(year, monthNumber, 0).getDate();
  var daysInFirstWeek = 7-firstDay
  var remainingDays = daysInMonth - daysInFirstWeek
  var daysInLastWeek = remainingDays % 7
  var fullWeeks = (daysInMonth - daysInFirstWeek - daysInLastWeek)/7

  return [daysInFirstWeek, fullWeeks, daysInLastWeek]
}

function getRangeForMonth(originRow, originCol, monthNumber, year){
  var monthStruct = getMonthStruct(monthNumber, year)
  var cellsInMonth = []

  var emptyDays = 7 - monthStruct[0]
  var fullWeeks = monthStruct[1]
  var daysinLastWeek = monthStruct[2]

  for (var i = originCol + 2*(emptyDays); i < originCol + 2*(7); i = i + 2){
    cellsInMonth.push([originRow, i])
  }
  for (var i = 1; i <= fullWeeks; i++){
    for (var j = 0; j < 7*2; j = j + 2){
      cellsInMonth.push([originRow+i, originCol+j])
    }
  }
  for (var i = originCol; i < originCol + 2*(daysinLastWeek); i=i+2){
    cellsInMonth.push([originRow+fullWeeks+1, i])
  }

  return cellsInMonth

}

function createGrid(sheet, originRow, originCol, monthNumber, year){
  var monthRange = getRangeForMonth(originRow, originCol, monthNumber, year)
  var day = 1
  sheet.getRange(originRow-1, originCol).setValue(months[monthNumber-1])
  for (var i = 0; i < monthRange.length; i++){
    var cellWithNumber = sheet.getRange(monthRange[i][0], monthRange[i][1])
    var cellWithCheckbox = sheet.getRange(monthRange[i][0], monthRange[i][1]+1)
    cellWithNumber.setValue(day)
    cellWithCheckbox.insertCheckboxes()

    cellWithNumber.setBorder(true, true, true, false, false, false)
    cellWithCheckbox.setBorder(true, false, true, true, false, false)
    
    day++
  }

  for (var i = originCol; i < originCol + 14; i++){
    sheet.setColumnWidth(i, 21)
  }
}

function addDaysPerWeekCount(sheet, originRow, originCol, monthStruct, averageTimesPerWeekRange, prevMonth){
  
  var daysInFirstWeek = monthStruct[0]
  var fullWeeks = monthStruct[1]
  var daysInLastWeek = monthStruct[2]
  
  var firstCellInRow = [originRow, originCol + 2*(7-daysInFirstWeek)]
  var lastCellInRow = [originRow, originCol + 13]


  sheet.getRange(lastCellInRow[0]-1,lastCellInRow[1]+1).setValue("Days per week")
  
  var prevMonthSumRange = ""
  if (prevMonth != null) {
    prevMonthSumRange = "+countif(" + prevMonth + ", TRUE)" 
  }

  var fRow = lastCellInRow[0]
  var fCol = lastCellInRow[1]+1

  for (var i = 0; i <= fullWeeks; i++){
    var sumRange = convertToCellReference(firstCellInRow[0], firstCellInRow[1])
    sumRange += ":"
    sumRange += convertToCellReference(lastCellInRow[0], lastCellInRow[1])
    if (i == 0){
      sheet.getRange(lastCellInRow[0], lastCellInRow[1] + 1).setFormula("=countif("+sumRange+",TRUE)"+prevMonthSumRange)
    }
    else{
      sheet.getRange(lastCellInRow[0], lastCellInRow[1] + 1).setFormula("=countif("+sumRange+",TRUE)")
    }
    firstCellInRow = [firstCellInRow[0]+1, originCol]
    lastCellInRow = [lastCellInRow[0]+1, originCol+13]
  }

  var range1 = convertToCellReference(fRow, fCol) + ":" + convertToCellReference(fRow+4, fCol)
  var range2 = convertToCellReference(fRow, fCol - 15) + ":" + convertToCellReference(fRow+4,fCol-15)
  averageTimesPerWeekRange.push([range1, range2])

  if(daysInLastWeek > 0){
    var sumRange = convertToCellReference(firstCellInRow[0], firstCellInRow[1])
    sumRange += ":"
    sumRange += convertToCellReference(lastCellInRow[0], originCol+2*daysInLastWeek-1)
    return {sumRange, averageTimesPerWeekRange}
  }
  var res = null
  return {res, averageTimesPerWeekRange}
}

function addWeekNumber(sheet, originRow, originCol, monthStruct, monthNumber, weekNumberRange, year){
  var daysInFirstWeek = monthStruct[0]
  var fullWeeks = monthStruct[1]
  var lastWeekOfTheYear;

  var day = new Date(year, monthNumber-1, 1)
  var weekNumber = Utilities.formatDate(day, "GMT", "w");
  sheet.getRange(originRow, originCol-1).setValue(weekNumber).setFontSize(8)


  for (var i = 0; i < fullWeeks; i++){
    var dd = 7*i+daysInFirstWeek+1
    var day = new Date(year, monthNumber-1, dd)
    var weekNumber = Utilities.formatDate(day, "GMT", "w")
    lastWeekOfTheYear = weekNumber;
    sheet.getRange(originRow+1+i, originCol-1).setValue(weekNumber).setFontSize(8)
    weekNumberRange.push([originRow+1+i], originCol-1)
  }
  return weekNumberRange, lastWeekOfTheYear
  
}

function addHeaderFormulas(sheet, goal, averageTimesPerWeekRange, lastWeekOfTheYear){
  var yearAverageFormula = "=IFERROR(AVERAGE(FILTER("
  var averages = "FLATTEN({"
  var averagesSum = "=SUM("
  var weeks = "FLATTEN({"
  sheet.getRange(2, 17).setValue("Average times per week:")
  for (var i = 0; i < averageTimesPerWeekRange.length; i++){
    averages += averageTimesPerWeekRange[i][0]
    averagesSum += averageTimesPerWeekRange[i][0]
    weeks += averageTimesPerWeekRange[i][1]

    if (i != averageTimesPerWeekRange.length - 1){
      averages += ", "
      averagesSum += ", "
      weeks += ", "
    }
  }
  averages += "})"
  averagesSum += ")"
  weeks += "})"
  yearAverageFormula += averages + ", " + weeks + '< WEEKNUM(TODAY()))), "0")'

  sheet.setFrozenRows(6)
  sheet.getRange(2,24).setValue(yearAverageFormula)
  sheet.getRange(1,26).setValue("Goal");sheet.getRange(2,26).setValue(goal);
  sheet.getRange(3,17).setValue("Accomplished:")
  
  goalTotalDays = goal*lastWeekOfTheYear;

  sheet.getRange("U3:AA6").merge().setValue(averagesSum+"/"+goalTotalDays+"*100").setFontSize(56).setFontWeight("bold").setNumberFormat("#,##0.0")
  sheet.getRange("AB3:AB6").merge().setValue("%").setFontSize(16).setFontWeight("bold")

  sheet.getRange("F1:L6").merge().setValue(averagesSum).setFontSize(79).setFontWeight("bold")
  sheet.getRange("M5:O5").merge().setValue("/"+goalTotalDays).setFontSize(11).setFontWeight("bold")
}

function convertToCellReference(row, column) {
  // Convert column number to letter
  var columnLetter = String.fromCharCode(65 + (column - 1) % 26);
  if (column > 26) {
    columnLetter = String.fromCharCode(64 + Math.floor((column - 1) / 26)) + columnLetter;
  }

  // Combine row and column to form cell reference
  var cellReference = columnLetter + row;
  
  return cellReference;
}

















