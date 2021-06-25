const FOLDER_ID = "16-dser3Xd-kBGllite8tfkFdhCDCK5nL"
const REQUIRED_HEADER = [
  "STEP", 
  "TESTER ACTION", 
  "EXPECTED RESULT", 
  "STATUS", 
  "ISSUE #", 
  "RESULTS/NOTES"
]

const STATUS_BLANK = "◯ Pass\n◯ Fail\n◯ Blocked"

const arrayCompare = (arr1, arr2) => (
  Array.isArray(arr1)
  && Array.isArray(arr2) 
  && arr1.length === arr2.length 
  && arr1.every((v, i) => arr2[i] === v)
)

function clearComments() {
  const folder = DriveApp.getFolderById(FOLDER_ID).getFiles()
  while(folder.hasNext()) {
    const currentFile = folder.next()
    const doc = DocumentApp.openById(currentFile.getId())
    const docTables = doc.getBody().getTables()
    for(let table of docTables) {
      const headerRow = table.getRow(0)
      const tableHeader = [...Array(headerRow.getNumCells())].map((_, i) => headerRow.getCell(i).getText())      
      if(arrayCompare(tableHeader, REQUIRED_HEADER)) {
        for(let i = 1; i < table.getNumRows(); i++) {
          const currentRow = table.getRow(i);
          for(let j = 3; j <= 5; j++) {
            currentRow.getCell(j).setText(j === 3 ? STATUS_BLANK : '')
          }
        }
      }
    }
  }
}
