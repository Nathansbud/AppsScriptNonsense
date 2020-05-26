var weekMS = 604800000

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu("Update Data").addItem("Update Revision Cells", "getRevisions").addToUi()
}


function getRevisions() {
  let keys = ["date", "name"]
  let revisionItems = Drive.Revisions.list(SpreadsheetApp.getActive().getId()).items.reverse()
  let date = new Date()
  let revision = []
  
  for(r of revisionItems) {
    if(date - Date.parse(r['modifiedDate']) <= weekMS) {
       revision.push({"date":r['modifiedDate'], "name":r['lastModifyingUser']['displayName']})
    } else {
       break
    }
  }
      
  let currentSheet = SpreadsheetApp.openById(SpreadsheetApp.getActive().getId()).getActiveSheet()
  currentSheet.getRange(1, 1, revision.length, keys.length).setValues(revision.map(v => [v['date'], v['name']]))
}