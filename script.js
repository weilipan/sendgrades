function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu("成績處理")
    .addItem("產生成績表", "generateGradeSheet")
    .addItem("寄送成績", "sendGrades")
    .addToUi();
}

// 產生成績表
function generateGradeSheet() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var ui = SpreadsheetApp.getUi();
  
  // 輸入教師信箱
  var email;
  while (true) {
    var teacherEmail = ui.prompt("請輸入教師信箱", "請確保信箱格式正確，否則無法寄送成績！", ui.ButtonSet.OK_CANCEL);
    if (teacherEmail.getSelectedButton() == ui.Button.CANCEL) return;
    email = teacherEmail.getResponseText().trim();
    if (email.match(/^[^\s@]+@[^\s@]+\.[^\s@]+$/)) break;
    ui.alert("錯誤", "請輸入有效的電子郵件地址！", ui.ButtonSet.OK);
  }

  // 輸入課程名稱
  var course;
  while (true) {
    var courseName = ui.prompt("請輸入課程名稱", "此名稱將用於成績通知郵件", ui.ButtonSet.OK_CANCEL);
    if (courseName.getSelectedButton() == ui.Button.CANCEL) return;
    course = courseName.getResponseText().trim();
    if (course !== "") break;
    ui.alert("錯誤", "課程名稱不可為空！", ui.ButtonSet.OK);
  }

  // 清空試算表並建立標題列
  sheet.clear();
  sheet.appendRow([email, course]); // 教師信箱和課程名稱
  sheet.appendRow(["信箱", "學號", "班級", "座號", "姓名", "總成績"]); // 標題列
  ui.alert("成功", "成績表已建立，請輸入學生成績。", ui.ButtonSet.OK);
}

// 寄送成績
function sendGrades() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var data = sheet.getDataRange().getValues();
  var ui = SpreadsheetApp.getUi();
  
  if (data.length < 3) {
    ui.alert("錯誤", "試算表無有效資料，請先輸入學生成績！", ui.ButtonSet.OK);
    return;
  }
  
  var teacherEmail = data[0][0].trim();
  var courseName = data[0][1].trim();

  // 檢查教師信箱
  while (!teacherEmail.match(/^[^\s@]+@[^\s@]+\.[^\s@]+$/)) {
    var teacherEmailPrompt = ui.prompt("錯誤", "教師信箱格式不正確，請重新輸入正確的電子郵件：", ui.ButtonSet.OK_CANCEL);
    if (teacherEmailPrompt.getSelectedButton() == ui.Button.CANCEL) return;
    teacherEmail = teacherEmailPrompt.getResponseText().trim();
    sheet.getRange(1, 1).setValue(teacherEmail);
  }

  // 檢查課程名稱
  while (courseName === "") {
    var courseNamePrompt = ui.prompt("錯誤", "課程名稱未設定，請重新輸入：", ui.ButtonSet.OK_CANCEL);
    if (courseNamePrompt.getSelectedButton() == ui.Button.CANCEL) return;
    courseName = courseNamePrompt.getResponseText().trim();
    sheet.getRange(1, 2).setValue(courseName);
  }

  var headers = data[1]; // 標題列
  var date = new Date().toISOString().split("T")[0]; // YYYY-MM-DD
  var emailCount = 0; // 成功寄送計數
  
  for (var i = 2; i < data.length; i++) {
    var studentData = data[i];
    var studentEmail = studentData[0].trim();
    var studentName = studentData[4].trim();
    
    if (!studentEmail.match(/^[^\s@]+@[^\s@]+\.[^\s@]+$/)) {
      Logger.log("跳過無效信箱: " + studentEmail);
      continue;
    }
    
    // 建立成績內容
    var gradeDetails = "";
    for (var j = 5; j < studentData.length; j++) {
      if (headers[j] && studentData[j] !== "") {
        gradeDetails += `<b>${headers[j]}:</b> ${studentData[j]}<br>`;
      }
    }

    var subject = `[${courseName}]課程[${date}]成績試算-[${studentName}]同學`;
    var body = `
      <div style="font-family: Arial, sans-serif; line-height: 1.6;">
        <p><b>寄件者:</b> ${teacherEmail}</p>
        <p>親愛的 <b>${studentName}</b> 同學您好，</p>
        <p>您選修的 <b>${courseName}</b> 課程，目前 ${date} 成績結算如下：</p>
        <div style="background-color:#f8f8f8; padding:10px; border-left: 4px solid #007bff;">
          ${gradeDetails}
        </div>
        <p>如有疑問請於上課時與教師確認，感謝。</p>
        <p style="color:gray;"><i>最終成績請以校務行政系統為準。</i></p>
      </div>
    `;

    MailApp.sendEmail({
      to: studentEmail,
      cc: teacherEmail,
      subject: subject,
      htmlBody: body
    });

    emailCount++; // 計數成功寄出的郵件
  }
  
  ui.alert("寄送完成", `總共成功寄送 ${emailCount} 封郵件！`, ui.ButtonSet.OK);
}
