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
