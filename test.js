function test_readAttendanceSubject() {
  const ss = SpreadsheetApp.openByUrl('https://docs.google.com/spreadsheets/d/1PSS_v3A5h0srfkSJXeteHQbpXFkpZYDIMlkuuJk3sCM/edit')
  const result = readAttendanceSubject(ss)

  Logger.log(result)
}


function test_isExistTitle()
{
  const ss = SpreadsheetApp.openByUrl('https://docs.google.com/spreadsheets/d/1PSS_v3A5h0srfkSJXeteHQbpXFkpZYDIMlkuuJk3sCM/edit')
  const subjects = readAttendanceSubject(ss)

  const subject = subjects[0]
  const sheet = ss.getSheetByName('鳴り物')
  const title = `${subject['type']} ${subject['title']}`
  const result = isExistTitle(sheet, title, subject['date'])

  Logger.log(result)
}


function test_getSubjects()
{
  const ss = SpreadsheetApp.openByUrl('https://docs.google.com/spreadsheets/d/1PSS_v3A5h0srfkSJXeteHQbpXFkpZYDIMlkuuJk3sCM/edit')
  const sheet = ss.getSheetByName('フォームの回答 2')
  const range = sheet.getRange('A2')

  const result = getSubjects(sheet, range)
  Logger.log(result)
}


function test_findSubjectColumnIndex()
{
  const ss = SpreadsheetApp.openByUrl('https://docs.google.com/spreadsheets/d/1PSS_v3A5h0srfkSJXeteHQbpXFkpZYDIMlkuuJk3sCM/edit')
  const sheet = ss.getSheetByName('鳴り物')

  const subject = {'title':'氏家', 'date': '2024/07/21'}
  const col_index = findSubjectColumnIndex(sheet, subject)

  Logger.log(`col index:${col_index}`)
}

function test_makeMgsAttendancesOnePart()
{
  const ss = SpreadsheetApp.openByUrl('https://docs.google.com/spreadsheets/d/1PSS_v3A5h0srfkSJXeteHQbpXFkpZYDIMlkuuJk3sCM/edit')
  const sheet = ss.getSheetByName('鳴り物')

  const subject = {'title':'氏家', 'date': '2024/07/21'}

  const col =  findSubjectColumnIndex(sheet, subject)
  const msg = makeMgsAttendancesOnePart(sheet, col)

  Logger.log(msg)
}


function test_makeMgsAttendancesAllParts()
{
  const subject = {'title':'氏家', 'date': '2024/07/21'}
  const msg =  makeMgsAttendancesAllParts(subject) 
  Logger.log(msg)  
}