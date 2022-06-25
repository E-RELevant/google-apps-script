/**
 * 
 * © 2021. All Rights Reserved.
 * Erel Adoni
 *
 */
function registerNewEditResponseURLTrigger() {
  // check if an existing trigger is set
  var existingTriggerId = PropertiesService.getUserProperties().getProperty('onFormSubmitTriggerID')
  if (existingTriggerId) {
    var foundExistingTrigger = false
    ScriptApp.getProjectTriggers().forEach(function (trigger) {
      if (trigger.getUniqueId() === existingTriggerId) {
        foundExistingTrigger = true
      }
    })
    if (foundExistingTrigger) {
      return
    }
  }
  var trigger = ScriptApp.newTrigger('onFormSubmitEvent')
    .forSpreadsheet(SpreadsheetApp.getActive())
    .onFormSubmit()
    .create()
  PropertiesService.getUserProperties().setProperty('onFormSubmitTriggerID', trigger.getUniqueId())
}
function getTimestampColumn(sheet) {
  for (var i = 1; i <= sheet.getLastColumn(); i += 1) {
    if (sheet.getRange(1, i).getValue() === 'Timestamp') {
      return i
    }
  }
  return 1
}
function getFormResponseEditUrlColumn(sheet) {
  var form = FormApp.openByUrl(sheet.getFormUrl())
  for (var i = 1; i <= sheet.getLastColumn(); i += 1) {
    if (sheet.getRange(1, i).getValue() === 'קישור לעריכת נתוני התגובה') {
      return i
    }
  }
  // get the last column at which the url can be placed.
  return Math.max(sheet.getLastColumn() + 1, form.getItems().length + 2)
}
/**
 * params: { sheet, form, formResponse, row }
 */
function addEditResponseURLToSheet(params) {
  if (!params.col) {
    params.col = getFormResponseEditUrlColumn(params.sheet)
  }
  var formResponseEditUrlRange = params.sheet.getRange(params.row, params.col)
  formResponseEditUrlRange.setValue(params.formResponse.getEditResponseUrl())
}
function onOpen() {
  var menu = [{ name: '❮/❯ Script: add links to edit form\'s comments', functionName: 'setupFormEditResponseURLs' }]
  SpreadsheetApp.getActive().addMenu('Custom', menu)
}
function setupFormEditResponseURLs() {
  var sheet = SpreadsheetApp.getActiveSheet()
  var spreadsheet = SpreadsheetApp.getActive()
  var formURL = sheet.getFormUrl()
  if (!formURL) {
    SpreadsheetApp.getUi().alert('⚠️⚠️⚠️\r\nThis spreadsheet is not linked to a Google form. Select a destination for the comments in your form to link to the spreadsheet.\r\nThe relevant sheet should be exposed if the comments sheet of the form is hidden.')
    return
  }
  var form = FormApp.openByUrl(formURL)
  // setup the header if not existed
  var headerFormEditResponse = sheet.getRange(1, getFormResponseEditUrlColumn(sheet))
  var title = headerFormEditResponse.getValue()
  if (!title) {
    headerFormEditResponse.setValue('Link to edit the response')
  }
  var timestampColumn = getTimestampColumn(sheet)
  var editResponseUrlColumn = getFormResponseEditUrlColumn(sheet)
  
  var timestampRange = sheet.getRange(2, timestampColumn, sheet.getLastRow() - 1, 1)
  var editResponseUrlRange = sheet.getRange(2, editResponseUrlColumn, sheet.getLastRow() - 1, 1)
  if (editResponseUrlRange) {
    var editResponseUrlValues = editResponseUrlRange.getValues()
    var timestampValues = timestampRange.getValues()
    for (var i = 0; i < editResponseUrlValues.length; i += 1) {
      var editResponseUrlValue = editResponseUrlValues[i][0]
      var timestampValue = timestampValues[i][0]
      if (editResponseUrlValue === '') {
        var timestamp = new Date(timestampValue)
        if (timestamp) {
          var formResponse = form.getResponses(timestamp)[0]
          editResponseUrlValues[i][0] = formResponse.getEditResponseUrl()
          var row = i + 2
          if (row % 10 === 0) {
            spreadsheet.toast('⟳ Proccessing lines ' + row + ' to ' + (row + 10))
            editResponseUrlRange.setValues(editResponseUrlValues)
            SpreadsheetApp.flush()
          }
        }
      }
    }
    
    editResponseUrlRange.setValues(editResponseUrlValues)
    SpreadsheetApp.flush()
  }
  registerNewEditResponseURLTrigger()
  
  SpreadsheetApp.getUi().alert('The script has successfully completed its work! In this worksheet, look for the column \'Link to edit the response\'. Future comments will be automatically synced with the link to edit the form.\r\n\r\nEnjoy.')
}
function onFormSubmitEvent(e) {
  var sheet = e.range.getSheet()
  var form = FormApp.openByUrl(sheet.getFormUrl())
  var formResponse = form.getResponses().pop()
  addEditResponseURLToSheet({
    sheet: sheet,
    form: form,
    formResponse: formResponse,
    row: e.range.getRow(),
  })
}