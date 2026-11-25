// 연결할 구글 스프레드시트 ID를 입력하세요.
// (스프레드시트 URL: https://docs.google.com/spreadsheets/d/여기가_ID/edit)
// 비워두면 스크립트에 바인딩된 스프레드시트 또는 새 스프레드시트를 사용합니다.
const SPREADSHEET_ID = '1ZfH0BQHMglmiDmtvib3rLss3VnBSeqOzTrzPYPAswfw';
const SHEET_NAME = '1125 lunch menu'; // 실제 탭 이름에 맞게 수정하세요.

function doGet(e) {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('11월25일 저녁 신청 받기')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
}

function getTargetSheet_() {
  let ss;
  if (SPREADSHEET_ID && SPREADSHEET_ID.trim() !== '') {
    ss = SpreadsheetApp.openById(SPREADSHEET_ID.trim());
  } else {
    ss = SpreadsheetApp.getActiveSpreadsheet();
  }

  if (!ss) {
    throw new Error('연결된 구글 스프레드시트를 찾을 수 없습니다. SPREADSHEET_ID를 설정해 주세요.');
  }

  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }

  const expectedHeaders = ['이름', '음식', '하고 싶은 말'];
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(expectedHeaders);
  } else {
    const firstRow = sheet.getRange(1, 1, 1, expectedHeaders.length).getValues()[0];
    const isEmptyHeader = firstRow.every(cell => cell === '');
    if (isEmptyHeader) {
      sheet.getRange(1, 1, 1, expectedHeaders.length).setValues([expectedHeaders]);
    }
  }

  return sheet;
}

function submitApplication(formData) {
  try {
    if (!formData) {
      return { success: false, message: '입력 데이터가 없습니다.' };
    }

    const name = String(formData.name || '').trim();
    const food = String(formData.food || '').trim();
    const message = String(formData.message || '').trim();

    if (!name || !food) {
      return { success: false, message: '이름과 음식을 입력해 주세요.' };
    }

    const lock = LockService.getScriptLock();
    lock.waitLock(10000);

    try {
      const sheet = getTargetSheet_();
      const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

      // 헤더 순서가 '이름', '음식', '하고 싶은 말'로 되어 있는지 확인 후 매핑
      const rowValues = new Array(headers.length).fill('');
      const mapping = {
        '이름': name,
        '음식': food,
        '하고 싶은 말': message
      };

      let mappedAny = false;
      headers.forEach((header, idx) => {
        const key = String(header).trim();
        if (Object.prototype.hasOwnProperty.call(mapping, key)) {
          rowValues[idx] = mapping[key];
          mappedAny = true;
        }
      });

      if (!mappedAny) {
        sheet.appendRow([name, food, message]);
      } else {
        sheet.appendRow(rowValues);
      }

      return { success: true, message: '답장 감사합니다.' };
    } finally {
      lock.releaseLock();
    }
  } catch (err) {
    return {
      success: false,
      message: '저장 중 오류가 발생했습니다: ' + err.message
    };
  }
}
