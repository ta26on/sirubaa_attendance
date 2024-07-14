function readAttendanceSubject(ss) {
  const sheet = ss.getSheetByName(FORM_NAME)

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
  const ss = SpreadsheetApp.openByUrl(SS_URL)
  const subjects = readAttendanceSubject(ss)

  subjects.forEach((subject) => {
    PARTS.forEach((part) => {
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
function getSubjectFromFormSheet() {
  const sheet = SpreadsheetApp.getActiveSheet()
  const range = sheet.getActiveRange()
  return getSubjects(sheet, range)
}

function getSubjects(sheet, range) {
  const row = range.getRow()
  ret = {}
  ret['title'] = sheet.getRange(row, COL_INDEX_TITLE).getValue()
  ret['date'] = sheet.getRange(row, COL_INDEX_DATE).getValue()
  return ret
}

function findSubjectColumnIndex(sheet, subject) {
  const range = sheet.getDataRange()
  for (let col = 1; col < range.getLastColumn(); col++) {
    const title = range.getCell(1, col).getValue()
    if (title.includes(subject['title'])) {
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

function makeMgsAttendancesOnePart(sheet, col) {
  let msg = ''
  const num_rows = sheet.getMaxRows()

  let attendances = {}

  RESPONSE_TYPE.forEach((type) => {
    msg += "\\n"
    attendances[type] = []
    for (let row = 5; row < num_rows; row++) {
      const answer = sheet.getRange(row, col).getValue()
      if (answer.includes(type)) {
        const member_name = sheet.getRange(row, 1).getValue()
        attendances[type].push(`${member_name} ${answer.replace(type, '')}`)
      }
    }
    if (attendances[type].length <= 0) { return }
    msg += `${type}: ${attendances[type].length}\\n`
    attendances[type].forEach(row => {
      msg += row + "\\n"
    })
  })
  return msg
}

function makeMgsAttendancesAllParts(subject) {
  const created = new Date()
  const created_str = Utilities.formatDate(created, 'JST', 'yyyy-MM-dd HH:mm:ss');
  let msg = `出力日時: ${created_str}\\n\\n`

  const subject_date = new Date(subject['date'])
  const subject_date_str = Utilities.formatDate(subject_date, 'JST', 'yyyy-MM-dd'); // todo 曜日

  msg += `${subject_date_str} ${subject['title']}`;

  const ss = SpreadsheetApp.openByUrl(SS_URL)

  PARTS.map(part => {
    const sheet = ss.getSheetByName(part)
    const col_index = findSubjectColumnIndex(sheet, subject)
    if (col_index != -1) {
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