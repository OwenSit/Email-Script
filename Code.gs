/**
 * Returns a rectangular grid of values in a given sheet.
 * @param {string} sheetName The name of the sheet.
 * @return {object[][]} A two-dimensional array of values in the sheet.
 */

// Acquire email addresses and marks 
function getData(sheetName) {
  var data = SpreadsheetApp.getActive().getSheetByName(sheetName).getDataRange().getValues();
  return data;
}

/**
 * Sends an email for each row.
 */
function sendEmails() {
  var sheetData = getData("Marks");
  var CourseInfo = getData("Course Info");
  var courseNo = CourseInfo[1][0]
  
  for (k = 1; k < sheetData.length; k++) {
    
    var lastName = sheetData[k][0].split(",")[0];
    var firstName = sheetData[k][0].split(",")[1];
    var email = sheetData[k][2];
    var class_list = sheetData[k][3]
    var EMAIL_SUBJECT = courseNo + " | Marks for" + firstName + " " + lastName; 
    var marks = [];
    var greet = {};
    var styles = {};
    styles.odd = "overflow:hidden;padding:2px 3px;vertical-align:bottom;border:1px solid rgb(204,204,204)";
    styles.even = "overflow:hidden;padding:2px 3px;vertical-align:bottom;background-color:rgb(221,242,240);text-align:right;border:1px solid rgb(204,204,204)"
    greet.intro1 = "Hello" + firstName + " " + lastName + ", classlist#: " + class_list;
    greet.intro2 = "Please see below you current marks in "  + courseNo + ":\n"
    greet.outro1 = "Please let me know if you have any questions."
    greet.outro2 = "Regards,"
    greet.outro3 = "Owen, TA"

    for (j = 4; j < sheetData[k].length; j++) {
      var mark = {};
      mark.markContent = sheetData[0][j];
      mark.markScore = sheetData[k][j]==='' ? 'N/A' : sheetData[k][j]; 
      marks.push(mark);
    }


    var htmlTemplate = HtmlService.createTemplateFromFile("Template.html");
    htmlTemplate.marks = marks;
    htmlTemplate.greet = greet;
    htmlTemplate.styles = styles;
    


    var htmlBody = htmlTemplate.evaluate().getContent();

    

    MailApp.sendEmail({
        to: email,
        subject: EMAIL_SUBJECT,
        htmlBody: htmlBody
      });
  
  }
}

