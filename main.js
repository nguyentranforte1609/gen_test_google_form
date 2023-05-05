function genTestForm() {
  var data = getDataFromSheet()
  var test_data = data[0]
  var spreadsheet_name = data[1]
  var form_id = createNewForm(test_data, spreadsheet_name)
  var folder_name = spreadsheet_name.replace('Gen_', 'Test_')
  // Please make sure there is only one table with the that name
  var folder_id = DriveApp.getFoldersByName(folder_name).next().getId() 
  moveFormToTestFolder(form_id, folder_id)
}

function getDataFromSheet() {
  // Get sheet
  var active_spreadsheet = SpreadsheetApp.getActiveSpreadsheet()
  var summary_sheet = active_spreadsheet.getSheetByName("test_summary");
  var question_sheet = active_spreadsheet.getSheetByName("all_questions");

  // Get current test summary
  var num_total_quetions = summary_sheet.getRange("B1").getValue();
  var num_tested_questions = summary_sheet.getRange("B2").getValue();
  var num_question_per_day = summary_sheet.getRange("B3").getValue();
  var start_row = 0
  var num_curr_tested_questions = 0

  // Get questions
  // Check if end of test
  if (num_tested_questions === num_total_quetions) { // End of test => restart
    start_row = 2
    num_curr_tested_questions = 10
  }
  if (num_tested_questions + num_question_per_day > num_total_quetions) {
    // Final test
    start_row = num_tested_questions + 2
    num_curr_tested_questions = num_total_quetions
  }
  if (num_tested_questions + num_question_per_day <= num_total_quetions) {
    start_row = num_tested_questions + 2
    num_curr_tested_questions = num_tested_questions + num_question_per_day
  }
  var data = question_sheet.getRange(start_row, 1, num_question_per_day, question_sheet.getLastColumn()).getValues();
  
  // Update test summary
  summary_sheet.getRange("B2").setValue(num_curr_tested_questions)
  return [data, active_spreadsheet.getName()]
}

function createNewForm(data, test_name) {
  var today_string = Utilities.formatDate(new Date(), "GMT+7", "_yyyyMMdd_hhmm")
  var form = FormApp.create('Test ' + test_name + today_string);
  form.setIsQuiz(true);
  form.setProgressBar(true);
  for(var i=0; i < data.length; i++) {
    var answers = data[i][2].split('|')
    var correct_answers = data[i][4].split('|')
    if (correct_answers.length === 1) { // Single Answer
      var item = form.addMultipleChoiceItem()
      item.setTitle(data[i][0] + ' \n' + data[i][1])
      item.setChoices(answers.map((answer) => item.createChoice(
        answer
        , answer.includes(correct_answers[0])
      )))
      item.setRequired(true)
      if (data[i][3] != '') {
        var feedback = FormApp.createFeedback()
        feedback.setText(data[i][3])
        item.setFeedbackForIncorrect(feedback.build())
      }
    }
    else { // Multiple Answer
      var item = form.addCheckboxItem()
      item.setTitle(data[i][0] + ' \n' + data[i][1])
      item.setChoices(answers.map((answer) => item.createChoice(
        answer
        , correct_answers.some((ca) => answer.includes(ca))
      )))
      item.setRequired(true)
      if (data[i][3] != '') {
        var feedback = FormApp.createFeedback()
        feedback.setText(data[i][3])
        item.setFeedbackForIncorrect(feedback.build())
      }
    }
  }
  return form.getId()
}

function moveFormToTestFolder(form_id, folder_id) {
  form_file = DriveApp.getFileById(form_id)
  dest_folder = DriveApp.getFolderById(folder_id)
  form_file.moveTo(dest_folder)
}
