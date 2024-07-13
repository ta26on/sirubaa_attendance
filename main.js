function readAttendanceSubject(ss) {
  const sheet = ss.getSheetByName('フォームの回答 2')

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
    row['deadline'] = values[i][4] == null ? '' : values[i][4]
    ret.push(row)
  }

  ret.sort((a, b) => { a['date'] < b['date'] ? 1 : -1 })

  return ret
}

function isExistTitle(sheet, title, date) {
  values = sheet.getRange('b1:j2').getValues()
  for (let col = 0; col < values[0].length; col++) {
    _title = values[0][col]
    _date = values[1][col]

    if (title != _title) {
      continue
    }
    if (date.toDateString() != _date.toDateString()) {
      continue
    }
    return true
  }
  return false
}

function addColumn() {
  const ss = SpreadsheetApp.openByUrl('https://docs.google.com/spreadsheets/d/1PSS_v3A5h0srfkSJXeteHQbpXFkpZYDIMlkuuJk3sCM/edit')

  const subjects = readAttendanceSubject(ss)

  const parts = ['鳴り物', '女踊り', '男踊り(青)', '女男踊り(助)', '浴衣', '子供(黄)']
  //const parts = ['鳴り物']

  subjects.forEach((subject) => {
    parts.forEach((part) => {
      addColumnToSheet(ss.getSheetByName(part), subject)
    })
  })
}

function addColumnToSheet(sheet, subject) {
  const title = `${subject['type']} ${subject['title']}`

  if (isExistTitle(sheet, title, subject['date']) == true) { return }

  sheet.insertColumnBefore(2)

  sheet.getRange('b1').setValue(title)
  sheet.getRange('b2').setValue(subject['date'])
  sheet.getRange('b3').setValue(subject['deadline'])
}

//　
function getSubjectFromFormSheet()
{
  const sheet = SpreadsheetApp.getActiveSheet()
  const range = sheet.getActiveRange()
  return getSubjects(sheet, range)
}

function getSubjects(sheet, range)
{
  const row = range.getRow()
  const col_title = 3
  const col_date = 4
  ret = {}
  ret['title'] = sheet.getRange(row, col_title).getValue()
  ret['date'] = sheet.getRange(row, col_date).getValue()
  return ret
}

function findSubjectColumnIndex(sheet, subject)
{
  const range = sheet.getDataRange()
  for (let col=1; col < range.getLastColumn(); col++)
  {
    const title = range.getCell(1,col).getValue()
    if (title.includes(subject['title']))
    {
      return col
    }
  }
  return -1
}

function showAttendances() {
  const sheet = SpreadsheetApp.getActiveSheet()
  const col = sheet.getActiveRange().getColumn()
  let msg = makeMgsAttendancesOnePart(sheet, col)
  Browser.msgBox(msg)
}

function makeMgsAttendancesOnePart(sheet, col) 
{
  let msg = ''
//  const created = new Date()
//  const created_str = Utilities.formatDate(created, 'JST', 'yyyy-MM-dd HH:mm:ss');
//  msg += `出力日時: ${created_str}`

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
        attendances[type].push(`${member_name} ${answer.replace(type, '')}`)
      }
    }
    if (attendances[type].length <= 0){ return }
    msg += `${type}: ${attendances[type].length}\\n`
    attendances[type].forEach(row => {
      msg += row + "\\n"
    })
  })
  return msg
}

function makeMgsAttendancesAllParts(subject) 
{
  const created = new Date()
  const created_str = Utilities.formatDate(created, 'JST', 'yyyy-MM-dd HH:mm:ss');
  let msg = `出力日時: ${created_str}\\n\\n`

  const subject_date = new Date(subject['date'])
  const subject_date_str = Utilities.formatDate(subject_date, 'JST', 'yyyy-MM-dd'); // todo 曜日

  msg += `${subject_date_str} ${subject['title']}`;

  const ss = SpreadsheetApp.openByUrl('https://docs.google.com/spreadsheets/d/1PSS_v3A5h0srfkSJXeteHQbpXFkpZYDIMlkuuJk3sCM/edit')

  const parts = ['鳴り物','女踊り', '男踊り(青)', '女男踊り(助)', '浴衣','子供(黄)']
  parts.map( part => {
    const sheet = ss.getSheetByName(part)
    const col_index = findSubjectColumnIndex(sheet, subject)
    if (col_index != -1)
    {
      msg += `\\n# ${part}`
      msg += makeMgsAttendancesOnePart(sheet, col_index)
    }
  })

  return msg
}

function showAttendancesAllParts() {
  const subject = getSubjectFromFormSheet()
  const msg = makeMgsAttendancesAllParts(subject) 
  Browser.msgBox(msg)
}

function onOpen() {
  const customMenu = SpreadsheetApp.getUi()
  customMenu.createMenu('しるばあ機能')
    .addItem('列追加', 'addColumn')
    .addItem('集計(パート毎)', 'showAttendances')
    .addItem('集計(全パート)', 'showAttendancesAllParts')
    .addToUi()
}