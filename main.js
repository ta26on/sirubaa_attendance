function readAttendanceSubject(ss)
{
  const sheet = ss.getSheetByName( 'フォームの回答 2' )

  // This represents ALL the data
  var range = sheet.getDataRange();
  var values = range.getValues();

  var ret = []
  // This logs the spreadsheet in CSV format with a trailing comma
  for (let i = 1; i < values.length; i++) {

    var row = {}
    row['type'] = values[i][1]
    row['title'] = values[i][2]
    row['date'] = values[i][3]
    row['deadline'] = values[i][4] == null ? '': values[i][4]
    ret.push(row)
  }

  ret.sort((a,b)=>{ a['date'] < b['date'] ? 1 : -1})

  return ret
}

function isExistTitle(sheet, title, date)
{
  values = sheet.getRange('b1:j2').getValues()
  for (let col=0; col < values[0].length; col++)
  {
    _title = values[0][col]
    _date = values[1][col]

    if (title != _title){
      continue
    }
    if(date.toDateString() != _date.toDateString())
    {
      continue
    }
      return true
  } 
  return false
}

function addColumn() {
    const ss = SpreadsheetApp.openByUrl('https://docs.google.com/spreadsheets/d/1PSS_v3A5h0srfkSJXeteHQbpXFkpZYDIMlkuuJk3sCM/edit')

    const subjects = readAttendanceSubject(ss)

    const parts = ['鳴り物', '女踊り', '男踊り(青)', '女男踊り(助)','浴衣','子供(黄)']
    //const parts = ['鳴り物']

    subjects.forEach((subject) => 
    {
      parts.forEach((part) => {
        addColumnToSheet(ss.getSheetByName(part), subject)
      })
    })
}

function addColumnToSheet(sheet, subject) {
  const title = `${subject['type']} ${subject['title']}`

  if (isExistTitle(sheet, title, subject['date'])==true){return}

  sheet.insertColumnBefore(2)

  sheet.getRange('b1').setValue(title)
  sheet.getRange('b2').setValue(subject['date'])
  sheet.getRange('b3').setValue(subject['deadline'])
}

function showAttendances() {
  const sheet = SpreadsheetApp.getActiveSheet()
  const col = sheet.getActiveRange().getColumn()
  let msg = sheet.getRange(1, col).getValue() + "\\n"

  const num_rows = sheet.getMaxRows()

  let attendances = {}

  const types = ['◯', '△']

  types.forEach((type) => {
    msg += "\\n"
    attendances[type] = []
    for (let row = 5; row < num_rows; row++) {
      const answer = sheet.getRange(row, col).getValue()
      if (answer.includes(type)) {
        const member_name = sheet.getRange(row, 1).getValue()
        attendances[type].push(`${member_name} ${answer.replace(type,'')}`)
      }
    }
    msg += `${type}: ${attendances[type].length}\\n`
    attendances[type].forEach(row => {
      msg += row + "\\n"
    })
  })

  Browser.msgBox(msg)
}

function onOpen() {
  const customMenu = SpreadsheetApp.getUi()
  customMenu.createMenu('しるばあ機能')  
      .addItem('列追加', 'addColumn')
      .addItem('集計', 'showAttendances')
      .addToUi()
}