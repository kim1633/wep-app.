const SHEET_ID = '1atFAgIsJMH5a59F2hjfDyN3Yo6ZpN8EkIYTqaTPYhu0';
const MAIN_SHEET = '출력결과';

function doGet(e) {
  const page = e && e.parameter && e.parameter.page;

  if (page === 'staff') {
    // QR 스캔으로 넘어온 쿠폰코드 (없으면 빈 문자열)
    const code = (e.parameter.code || '').trim();
    const tmpl = HtmlService.createTemplateFromFile('staff');
    tmpl.preloadCode = code; // staff.html로 코드 전달
    return tmpl.evaluate()
      .setTitle('복주걱 직원 확인')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('복주걱 10% 할인 쿠폰')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// 배포 URL 반환 (QR 생성용) ★ 신규
function getDeployUrl() {
  return ScriptApp.getService().getUrl();
}

// 쿠폰 발급 및 중복 체크
function processCouponRequest(clientId, ageGroup, gender, agreed, fingerprint) {
  if (!agreed) return { success: false, error: "개인정보 동의가 필요합니다." };

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName(MAIN_SHEET);
    const data = sheet.getDataRange().getValues();

    // 1차: 기존 세션ID로 검색 (기존 로직 유지)
    for (let i = 1; i < data.length; i++) {
      if (data[i][1] === clientId && data[i][6]) {
        return {
          success: true,
          code: data[i][6],
          date: Utilities.formatDate(data[i][0], "GMT+9", "yyyy-MM-dd"),
          isExisting: true,
          isUsed: !!data[i][7]
        };
      }
    }

    // 2차: 핑거프린트로 검색 (캐시 삭제 대응) ★ 신규
    if (fingerprint) {
      for (let i = 1; i < data.length; i++) {
        if (data[i][9] === fingerprint && data[i][6]) {
          return {
            success: true,
            code: data[i][6],
            date: Utilities.formatDate(data[i][0], "GMT+9", "yyyy-MM-dd"),
            isExisting: true,
            isUsed: !!data[i][7]
          };
        }
      }
    }

    // 신규 발급
    const couponCode = '복주걱-' + Utilities.getUuid().split('-')[0].toUpperCase();
    const now = new Date();
    // J열(index 9)에 fingerprint 추가, 기존 컬럼 순서 완전 유지
    sheet.appendRow([now, clientId, ageGroup, gender, 'TRUE', 'TRUE', couponCode, '', '', fingerprint || '']);

    return {
      success: true,
      code: couponCode,
      date: Utilities.formatDate(now, "GMT+9", "yyyy-MM-dd"),
      isUsed: false
    };
  } catch (e) {
    return { success: false, error: "접속자가 많습니다. 잠시 후 다시 시도해주세요." };
  } finally {
    lock.releaseLock();
  }
}

// 쿠폰 상태 조회
function checkCouponStatus(couponCode) {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(MAIN_SHEET);
  const data = sheet.getDataRange().getValues();
  for (let i = data.length - 1; i >= 1; i--) {
    if (data[i][6] === couponCode) {
      return {
        found: true,
        code: data[i][6],
        age: data[i][2],
        gender: data[i][3],
        issuedDate: Utilities.formatDate(data[i][0], "GMT+9", "yyyy-MM-dd HH:mm"),
        isUsed: !!data[i][7],
        usedDate: data[i][7] ? Utilities.formatDate(data[i][7], "GMT+9", "yyyy-MM-dd HH:mm") : null
      };
    }
  }
  return { found: false };
}

// 쿠폰 사용완료 처리
function redeemCouponByCode(couponCode) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(MAIN_SHEET);
    const data = sheet.getDataRange().getValues();
    for (let i = data.length - 1; i >= 1; i--) {
      if (data[i][6] === couponCode) {
        if (data[i][7]) return { success: false, error: '이미 사용 완료된 쿠폰입니다.' };
        sheet.getRange(i + 1, 8).setValue(new Date());
        return { success: true };
      }
    }
    return { success: false, error: '유효하지 않은 쿠폰입니다.' };
  } finally {
    lock.releaseLock();
  }
}

// 직원 비밀번호 검증
function verifyStaffPassword(pw) {
  const stored = PropertiesService.getScriptProperties().getProperty('STAFF_PW') || '1234';
  return pw === stored;
}
