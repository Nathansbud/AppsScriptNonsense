const sheet = SpreadsheetApp.getActive()
const SKILL_FORM_ID = "1JKYwNvppN6zXYVwOm-wZEbw6tR7TWd4Un225xviJYOM"
const MATRIX_TAB = "Demo Matrix"

function summarizeRanges() {
  const matrixTab = sheet.getSheetByName(MATRIX_TAB)
  const relevantRange = matrixTab.getRange(1, 1, 7, matrixTab.getLastColumn()).getValues()

  const sectionArr = relevantRange[0].slice(3)
  for(let i = 0; i < sectionArr.length; i++) {
    if(!sectionArr[i] && i > 0) sectionArr[i] = sectionArr[i - 1] 
  }

  const VALID_SECTIONS = Array.from(new Set(sectionArr)).map(k => [k, {}])

  const HEADERS = relevantRange[1].slice(3)
  const responses = relevantRange.slice(2).map(row => ({
    'name': row[0],
    'team': row[1],
    'email': row[2],
    'data': row.slice(3).reduce((acc, curr, i) => {
      if(curr) acc[sectionArr[i]][HEADERS[i].toLowerCase().trim()] = curr
      return acc
    }, Object.fromEntries(VALID_SECTIONS))
  }))

  return responses
}

function buildForm(results) {
  const skillForm = FormApp.openById(SKILL_FORM_ID)
  try {
    const formResponse = skillForm.createResponse()
    const items = skillForm.getItems();
    let prefilledResponse = formResponse.withItemResponse(
      items[0].asTextItem().createResponse(results.name)
    )

    for(let item of items) {
      const itemType = item.getType()
      const title = item.getTitle()
      const desc = item.getHelpText()
      if(itemType === FormApp.ItemType.GRID) {
        const gridItem = item.asGridItem()
        const gridValues = gridItem.getRows().map(v => {
          return results.data[title] ? (parseInt(results.data[title][v.toLowerCase().trim()] ?? '0').toString() || '0') : '0'
        })
        
        
        prefilledResponse = prefilledResponse.withItemResponse(gridItem.createResponse(gridValues))
      } else if(itemType == FormApp.ItemType.TEXT) {
        if(results.data[desc] && results.data[desc].other) {
          const textItem = item.asTextItem()
          prefilledResponse = prefilledResponse.withItemResponse(
            textItem.createResponse(results.data[desc].other || '')
          )
        } 
      }
    }
    return prefilledResponse.toPrefilledUrl()
  } catch(e) {
    console.log(e)
  }
}

function sendAllForms() {
    const employeeRows = summarizeRanges()
    for(let r of employeeRows.slice(0,-1)) {
      Logger.log(r)
      const formUrl = buildForm(r)
      const emailBody = constructEmail(r, formUrl)      
    }
}

function constructEmail({name, email}, formUrl) {
  const templateContent = HtmlService.createHtmlOutputFromFile("EmailTemplate.html").getContent()
  const emailTemplate = XmlService.parse(templateContent) 
  const root = emailTemplate.getRootElement()
  const emailBody = root.getChild('body').getChild('div')

  if(name) emailBody.getChild('h1').getChild('span').setText(name.split(' ')[0])
  emailBody.getChild('h2').getChild('a').setAttribute('href', formUrl)

  GmailApp.sendEmail(email, "Pre-filled Form Demo", "", {
    name: "Engineering Skills Survey",
    htmlBody: XmlService.getPrettyFormat().format(emailTemplate)
  })
}





