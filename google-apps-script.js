// Google Apps Script - 貼到 Google Apps Script 編輯器中
// 這個腳本會接收表單數據並寫入 Google Sheets

function doPost(e) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const data = JSON.parse(e.postData.contents);

    // 寫入數據（新增課程名稱欄位）
    sheet.appendRow([
      new Date().toLocaleString('zh-TW', { timeZone: 'Asia/Taipei' }),
      data.name,
      data.company,
      data.phone,
      data.email,
      data.participants,
      data.message || '',
      data.course_name || '未指定課程'
    ]);

    // 發送 Email 通知（可選）
    const emailBody = `
新的課程報名資料：

課程名稱：${data.course_name || '未指定課程'}
姓名：${data.name}
公司/部門：${data.company}
電話：${data.phone}
Email：${data.email}
預計人數：${data.participants}
備註：${data.message || '無'}

時間：${new Date().toLocaleString('zh-TW', { timeZone: 'Asia/Taipei' })}
    `;

    MailApp.sendEmail({
      to: 'nikeshoxmiles@gmail.com',
      subject: '【' + (data.course_name || '課程') + '】新報名通知 - ' + data.company,
      body: emailBody
    });

    return ContentService
      .createTextOutput(JSON.stringify({ success: true }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: error.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// 允許 GET 請求（用於測試和讀取數據）
function doGet(e) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const rows = data.slice(1);

    const result = rows.map(row => {
      const obj = {};
      headers.forEach((header, index) => {
        obj[header] = row[index];
      });
      return obj;
    });

    return ContentService
      .createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService
      .createTextOutput(JSON.stringify({ error: error.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
