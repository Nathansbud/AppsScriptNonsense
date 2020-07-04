//not sure if this is always right, but it seems to hold up pretty well...
//I figure the 2*arr.length accounts for the cell start/end points (or null terminator) and adding 1 is a 0-index thing? no clue really tho


const icsGreen = {color: {rgbColor: {red: 0.15, green: 0.62, blue: 0.57}}}
const icsText = {color: {rgbColor: {red: 0.14, green: 0.17, blue: 0.22}}}
const noColor = {color: {rgbColor: {}}}
const white = {color: {rgbColor: {red: 1, green: 1, blue: 1}}}
const transparent = {dashStyle: "SOLID", width: {magnitude: 0, unit: "PT"}, color: noColor} 

const indexOffset = (arr) => arr.reduce((acc, elem) => acc += elem.length, 0) + 2*arr.length+1

function makeTable(tableContent=[["Sr. Engineer", "100", "$20", "$2,000"], ["Skrt. Engineer", "100", "$20", "$2,000"]]) {
  const doc = DocumentApp.getActiveDocument()
  const body = doc.getBody()
  
  body.clear()
  
  const tableHeader = ["Role", "Hours", "Hourly Rate", "Total Price"]
  const tableFooter = ["", "", "Total Anticipated Budget", "$6,969"]
  
  let tableData = [tableHeader].concat(tableContent).concat([tableFooter])
  
  const table = doc.appendTable(tableData)
  const tableIdx = body.getChildIndex(table)
  const docId = doc.getId()
  doc.saveAndClose()
 
  //Use Docs API to style table, since App Script can't do it natively
  const tableElement = Docs.Documents.get(docId).body.content[tableIdx + 1]
  const tableStart = tableElement.startIndex
  const tableEnd = tableElement.endIndex
  
  const headerBorder = {dashStyle: "SOLID", width: {magnitude: 1, unit: "PT"}, color: icsGreen}
  const noBorder = {dashStyle: "SOLID", width: {magnitude: 1, unit: "PT"}, color: {}}
  
  let requests = {requests:[{
    updateTableCellStyle:{
      tableRange: {
        tableCellLocation: {
          tableStartLocation: {index: tableStart},
          rowIndex: 0, 
          columnIndex: 0
        }, 
        rowSpan: 1,
        columnSpan: tableHeader.length
      },
      tableCellStyle: {
        backgroundColor: icsGreen,
        contentAlignment: "MIDDLE" ,
        borderTop: headerBorder, borderLeft: headerBorder, borderRight: headerBorder, borderBottom: headerBorder
      },
      fields: "backgroundColor,borderTop,borderLeft,borderRight,borderBottom"
    }
  }, {
    updateTableCellStyle: {
      tableRange: {
        tableCellLocation: {
          tableStartLocation: {index: tableStart},
          rowIndex: 1,
          columnIndex: 0
        },
        rowSpan: tableContent.length,
        columnSpan: tableHeader.length
      },
      tableCellStyle: {
        borderBottom: {dashStyle: "SOLID", width: {magnitude: 1, unit: "PT"}, color: icsGreen},
        borderLeft: transparent,
        borderRight: transparent
      },
      fields: "borderBottom,borderLeft,borderRight"
    }
  }, {
    updateTableCellStyle: {
      tableRange: {
        tableCellLocation: {
          tableStartLocation: {index: tableStart},
          rowIndex: tableData.length - 1,
          columnIndex: 0
        },
        rowSpan: 1,
        columnSpan: tableFooter.length
      },
      tableCellStyle: {
        borderBottom: transparent,
        borderLeft: transparent,
        borderRight: transparent
      },
      fields: "borderBottom,borderLeft,borderRight"
    }
  }, {
    updateTextStyle: {
      fields:"foregroundColor",
      range: {
        startIndex: tableStart,
        endIndex: tableStart + indexOffset(tableHeader)
      },
      textStyle: {
        foregroundColor: white
      }
    }
  }, {
    updateTextStyle: {
      fields:"foregroundColor",
      range: {
        startIndex: tableStart + indexOffset(tableHeader),
        endIndex: tableEnd
      },
      textStyle: {
        foregroundColor: icsText
      }
    }
  }, {
    updateTextStyle: {
      fields: "fontSize",
      range: {
        startIndex: tableEnd - indexOffset(tableFooter),
        endIndex: tableEnd - indexOffset(tableFooter.slice(-1))
      },
      textStyle: {
        fontSize: {
          magnitude: 9,
          unit: "PT"
        }
      }
    } 
  }, {
    updateParagraphStyle: {
      paragraphStyle: {
        alignment: "CENTER"
      },
      fields: "alignment",
      range: {
        startIndex: tableStart,
        endIndex: tableStart + indexOffset(tableHeader)
      }
    } 
  }]}
  
  let io = tableStart + indexOffset(tableHeader) + 1
  tableContent.concat([tableFooter]).forEach(elem => {
    requests.requests.push({
      updateParagraphStyle: {
        paragraphStyle: {
          alignment: "END",
        },
        fields: "alignment",
        range: {
          startIndex: io + indexOffset(elem.slice(0, 1)),
          endIndex: io + indexOffset(elem)
        }
      }
    })
    io += indexOffset(elem)
  }) 
  
  Docs.Documents.batchUpdate(requests, docId);
}