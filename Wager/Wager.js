var dateColumn = 1
var startColumn = 2
var stopColumn = 3
var hoursColumn = 4
var summaryColumn = 9

function onEdit(e) {
  let sheet = SpreadsheetApp.getActiveSheet()
  if(e.range.offset(1, 1 - e.range.columnStart).getValue() == "TOTAL") {
    e.source.getActiveSheet().insertRowAfter(e.range.rowStart)
  }
  
  if(e.value == "=SUMMARIZE()") {
    sheet.getRange(e.range.getRow(), 1, 1, summaryColumn+1).setBackground("#000000").setFontColor("#FFFFFF").setFontWeight("bold").setFontSize(12).setHorizontalAlignment("center").setBorder(false, false, false, false, true, false, "white", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  }
  
  if(e.range.getColumn() == hoursColumn) {
    sheet.getRange(sheet.getLastRow(), e.range.getColumn()).setValue("") //this is inelegant but I want to force a refresh
    sheet.getRange(sheet.getLastRow(), e.range.getColumn()).setValue("=SUMT()")
  }
  
}
//sun 0-6 sat
var days = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"]

function SUMMARIZE() {
  let sheet = SpreadsheetApp.getActiveSheet()
  let currentRow = sheet.getActiveCell().getRow()  
  
  let values = []
  let weekSummary = {}
  days.forEach(d => {weekSummary[d] = 0})
 
  for(let i = currentRow - 1; i > 1; i--) {
    values = sheet.getRange(i, 1, 1, summaryColumn).getValues()[0]
    if(values[values.length - 1] == "Summary") {
      break
    }
    
    if(typeof values[dateColumn - 1].getMonth == 'function') {
      weekSummary[days[values[dateColumn - 1].getDay()]] += values[hoursColumn - 1].getHours() + values[hoursColumn -1].getMinutes() / 60
    }
  }
  
  let output = Object.entries(weekSummary).map(([k, v]) => `${k}: ${timeFormat(v)}`)
  let sundayValue = output.shift()
  output.push(sundayValue, "Total: " + timeFormat(sum(Object.values(weekSummary))), "Summary")
 
  return [output]
}

function timeFormat(dec) {
 let hours = Math.floor(dec) 
 let minutes = Math.round((dec - Math.floor(dec))*60)
 if(minutes < 10) minutes = "0"+minutes
 return hours +":"+minutes
}

function sum(obj) {
  return obj.reduce(function(a, b) {return a + b}, 0)
}

function SUMT() {
  let sheet = SpreadsheetApp.getActiveSheet()
  let column = sheet.getActiveCell().getColumn()
  let selectedValues = sheet.getRange(2, column, sheet.getLastRow() - 1, 1).getValues()
  Logger.log(selectedValues)
  totalSum = 0
  for(v of selectedValues) {
    if(typeof v[0].getMonth == 'function') totalSum += v[0].getHours() + v[0].getMinutes() / 60
  }
  return timeFormat(totalSum)
}
