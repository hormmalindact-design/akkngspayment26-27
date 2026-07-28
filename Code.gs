// ==========================================
// ⚙️ ផ្នែកកំណត់រចនាសម្ព័ន្ធ TELEGRAM
// ==========================================
var TELEGRAM_BOT_TOKEN = "8877919591:AAHy-g0du2GBVJx0sHisVFTIsD32NAd35qA"; 
var TELEGRAM_CHAT_ID = "-1004317236863";     

// 🌐 សម្រាប់ទទួលសំណើពី Vercel (API Endpoint)
// 🌐 សម្រាប់ទទួលសំណើ POST ពី Vercel (ដោះស្រាយបញ្ហា CORS)
function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.tryLock(10000);
  
  try {
    var data = JSON.parse(e.postData.contents);
    var action = data.action;
    var result = {};
    
    if (action === "addNewStudent") {
      result.studentId = addNewStudent(data.studentName, data.gender, data.studentClass, data.paymentType, data.amount, data.otherNote, data.schoolYear, data.fullYearFeeInput, data.paymentMethod, data.cashierName);
      result.status = "success";
    } else if (action === "updateStudentInfo") {
      result.message = updateStudentInfo(data.studentId, data.studentName, data.gender, data.studentClass, data.schoolYear, data.paymentType, data.otherNote, data.fullYearFeeInput, data.amountInput);
      result.status = "success";
    } else {
      result.status = "error";
      result.message = "Invalid action";
    }
    
    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
      
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({status: "error", message: err.toString()}))
      .setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

// 🌐 សម្រាប់ទទួលសំណើ GET ពី Vercel
// 🌐 សម្រាប់ទទួលសំណើ GET ពី Vercel (ដោះស្រាយបញ្ហា CORS ទាំងស្រុង)
// 🌐 សម្រាប់ទទួលសំណើ GET ពី Vercel
function doGet(e) {
  try {
    var action = e.parameter.action;
    var result = {};
    
    if (action === "getDashboardData") {
      result = getDashboardData();
    } else if (action === "getDailyClosingReport") {
      result = getDailyClosingReport();
    } else if (action === "getMonthlyClosingReport") {
      result = getMonthlyClosingReport();
    } else {
      result = { status: "error", message: "Invalid action" };
    }
    
    // បញ្ជូនទិន្នន័យចេញជា JSON (ភ្ជាប់ជាមួយ CORS Header ដោយស្វ័យប្រវត្តិ)
    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
      
  } catch (err) {
    // បើមាន Error ក៏ឱ្យលោតជា JSON ដែរ ដើម្បីកុំឱ្យជាប់ CORS 
    return ContentService.createTextOutput(JSON.stringify({ status: "error", message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
// ==========================================
// មុខងារគ្រឹះដដែល (Functions)
// ==========================================
function addNewStudent(studentName, gender, studentClass, paymentType, amount, otherNote, schoolYear, fullYearFeeInput, paymentMethod, cashierName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Students_Payment");
  var historySheet = ss.getSheetByName("Payment_History");
  
  var studentId = generateNextStudentId();
  var dateCreated = new Date();
  
  var actualAmount = Number(amount);
  var fullYearFee = Number(fullYearFeeInput);
  
  if (otherNote === "សិស្សក្រីក្រ (លើកលែង)" || otherNote === "សិស្សលើកលែង") {
    fullYearFee = 0;
    actualAmount = 0;
    paymentMethod = "លើកលែង (Exempted)";
  }
  
  var remainingBalance = fullYearFee - actualAmount;
  if (remainingBalance < 0) remainingBalance = 0;
  
  sheet.appendRow([
    studentId, studentName, studentClass, paymentType, actualAmount, 
    paymentMethod, dateCreated, "Paid", cashierName, otherNote, gender, schoolYear, fullYearFee, remainingBalance
  ]);
  
  var historyLog = (paymentType === "១ឆ្នាំពេញ") ? "បង់១ឆ្នាំពេញ" : "បង់ឆមាសទី១";
  if (historySheet) {
    historySheet.appendRow([studentId, studentName, dateCreated, historyLog, actualAmount, paymentMethod, cashierName]);
  }
  
  var dateStr = Utilities.formatDate(dateCreated, Session.getScriptTimeZone(), "dd/MM/yyyy hh:mm a");
  var telegramText = "🔔 <b><u>ជូនដំណឹងការចុះឈ្មោះ និងបង់ប្រាក់</u></b>\n" +
                     "--------------------------------------------------\n" +
                     "📝 <b>លេខវិក្កយបត្រ៖</b> " + studentId + "\n" +
                     "👤 <b>ឈ្មោះសិស្ស៖</b> " + studentName + "\n" +
                     "🏫 <b>ថ្នាក់រៀន៖</b> " + studentClass + "\n" +
                     "📦 <b>ដំណាក់កាលបង់៖</b> " + historyLog + "\n" +
                     "💰 <b>ទឹកប្រាក់ទទួលបាន៖</b> " + actualAmount.toLocaleString() + " KHR\n" +
                     "💵 <b>ប្រាក់ខ្វះ (ជំពាក់)៖</b> " + remainingBalance.toLocaleString() + " KHR\n" +
                     "💳 <b>វិធីសាស្ត្របង់៖</b> " + paymentMethod + "\n" +
                     "🧑‍💻 <b>បេឡាអ្នកទទួល៖</b> " + cashierName + "\n" +
                     "📅 <b>កាលបរិច្ឆេទ៖</b> " + dateStr + "\n" +
                     "--------------------------------------------------\n" +
                     "សាលាបឋមសិក្សាសម្តេចព្រះរាជអគ្គមហេសី នរោត្តម មុនីនាថ សីហនុ";
  sendTelegramMessage(telegramText);
  
  return studentId;
}

function updateStudentInfo(studentId, studentName, gender, studentClass, schoolYear, paymentType, otherNote, fullYearFeeInput, amountInput) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Students_Payment");
  var data = sheet.getDataRange().getValues();
  var searchId = String(studentId).trim();
  
  var actualAmount = Number(amountInput);
  var fullYearFee = Number(fullYearFeeInput);
  
  if (otherNote === "សិស្សក្រីក្រ (លើកលែង)" || otherNote === "សិស្សលើកលែង") {
    fullYearFee = 0;
    actualAmount = 0;
  }
  
  var remainingBalance = fullYearFee - actualAmount;
  if (remainingBalance < 0) remainingBalance = 0;
  
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === searchId) {
      var row = i + 1;
      sheet.getRange(row, 2).setValue(studentName);       
      sheet.getRange(row, 3).setValue(studentClass);      
      sheet.getRange(row, 4).setValue(paymentType);       
      sheet.getRange(row, 5).setValue(actualAmount);      
      sheet.getRange(row, 10).setValue(otherNote);        
      sheet.getRange(row, 11).setValue(gender);           
      sheet.getRange(row, 12).setValue(schoolYear);       
      sheet.getRange(row, 13).setValue(fullYearFee);      
      sheet.getRange(row, 14).setValue(remainingBalance); 
      
      if (remainingBalance > 0) {
         sheet.getRange(row, 8).setValue("Pending"); 
      } else {
         sheet.getRange(row, 8).setValue("Paid");
      }
      
      return "បានធ្វើបច្ចុប្បន្នភាពព័ត៌មាន និងទឹកប្រាក់របស់ " + studentName + " រួចរាល់!";
    }
  }
  return "រកមិនឃើញសិស្សដើម្បីកែប្រែឡើយ!";
}

function getDashboardData() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Students_Payment");
  if (!sheet) return { totalPaid: 0, totalFemale: 0, totalRevenue: "0 KHR", totalDiscounted: 0, totalExempted: 0, students: [] };
  var data = sheet.getDataRange().getValues();
  
  var totalPaidStudents = 0;
  var totalFemalePaid = 0;
  var totalRevenue = 0;
  var totalDiscounted = 0;
  var totalExempted = 0;
  var studentsList = [];
  
  for (var i = data.length - 1; i >= 1; i--) {
    var id = data[i][0];
    var name = data[i][1];
    var sClass = data[i][2];
    var type = data[i][3];
    var amount = Number(data[i][4]) || 0;
    var status = data[i][7];
    var other = data[i][9];
    var gender = data[i][10];
    var sYear = data[i][11] || "";
    var fullFee = Number(data[i][12]) || 0;
    var remaining = Number(data[i][13]) || 0;
    
    if (status === "Paid") {
      totalPaidStudents++;
      totalRevenue += amount;
      if (gender === "ស្រី") totalFemalePaid++;
    }
    
    if (other === "សិស្សបញ្ចុះតម្លៃ") {
      totalDiscounted++;
    } else if (other === "សិស្សក្រីក្រ (លើកលែង)" || other === "សិស្សលើកលែង") {
      totalExempted++;
    }
    
    if(studentsList.length < 15) {
      studentsList.push({
        id: id, name: name, gender: gender, class: sClass, type: type, 
        amount: amount, status: status, other: other, year: sYear, 
        fullFee: fullFee, remaining: remaining
      });
    }
  }
  
  return {
    totalPaid: totalPaidStudents,
    totalFemale: totalFemalePaid,
    totalRevenue: totalRevenue.toLocaleString() + " KHR",
    totalDiscounted: totalDiscounted,
    totalExempted: totalExempted,
    students: studentsList
  };
}

function getStudentById(studentId) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Students_Payment");
  var data = sheet.getDataRange().getValues();
  var searchId = String(studentId).trim();
  
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === searchId) {
      return {
        id: data[i][0], name: data[i][1], class: data[i][2], type: data[i][3],
        amount: data[i][4], method: data[i][5], status: data[i][7], other: data[i][9], 
        gender: data[i][10], year: data[i][11], fullFee: data[i][12], remaining: data[i][13] 
      };
    }
  }
  return null;
}

function collectSecondPayment(studentId, additionalAmount, paymentMethod, cashierName) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Students_Payment");
  var historySheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Payment_History");
  var data = sheet.getDataRange().getValues();
  var searchId = String(studentId).trim();
  
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === searchId) {
      var row = i + 1;
      var stuName = data[i][1];
      var stuClass = data[i][2];
      var currentAmount = Number(data[i][4]) || 0;
      var addAmt = Number(additionalAmount);
      var newAmount = currentAmount + addAmt; 
      
      sheet.getRange(row, 4).setValue("១ឆ្នាំពេញ");         
      sheet.getRange(row, 5).setValue(newAmount);         
      sheet.getRange(row, 6).setValue(paymentMethod);     
      sheet.getRange(row, 9).setValue(cashierName);       
      sheet.getRange(row, 14).setValue(0);                
      
      if(historySheet) {
        historySheet.appendRow([searchId, stuName, new Date(), "បង់បង្គ្រប់ (ឆមាសទី២)", addAmt, paymentMethod, cashierName]);
      }
      
      var dateStr = new Date().toLocaleString('en-GB', { hour12: true });
      var telegramText = "🔥 <b><u>ជូនដំណឹងការបង់ប្រាក់បង្គ្រប់ (លើកទី២)</u></b>\n" +
                         "--------------------------------------------------\n" +
                         "📝 <b>លេខវិក្កយបត្រ៖</b> " + searchId + "\n" +
                         "👤 <b>ឈ្មោះសិស្ស៖</b> " + stuName + "\n" +
                         "🏫 <b>ថ្នាក់រៀន៖</b> " + stuClass + "\n" +
                         "📦 <b>ដំណាក់កាលបង់៖</b> បង់បង្គ្រប់ (ឆមាសទី២) ✅\n" +
                         "💵 <b>ទឹកប្រាក់បង់បន្ថែម៖</b> " + addAmt.toLocaleString() + " KHR\n" +
                         "💰 <b>សរុបបានបង់ពេញ១ឆ្នាំ៖</b> " + newAmount.toLocaleString() + " KHR\n" +
                         "💵 <b>ប្រាក់ខ្វះ (ជំពាក់)៖</b> 0 KHR (បង់ដាច់)\n" +
                         "💳 <b>វិធីសាស្ត្របង់៖</b> " + paymentMethod + "\n" +
                         "🧑‍💻 <b>បេឡាអ្នកទទួល៖</b> " + cashierName + "\n" +
                         "📅 <b>កាលបរិច្ឆេទ៖</b> " + dateStr + "\n" +
                         "--------------------------------------------------\n" +
                         "សាលាបឋមសិក្សាសម្តេចព្រះរាជអគ្គមហេសី នរោត្តម មុនីនាថ សីហនុ";
      sendTelegramMessage(telegramText);
      
      return "ជោគជ័យ៖ បានទទួលប្រាក់បង្គ្រប់លើកទី២ ចំនួន " + addAmt.toLocaleString() + " KHR សម្រាប់សិស្ស " + stuName;
    }
  }
  return "រកមិនឃើញទិន្នន័យសិស្សឡើយ!";
}

function getDailyClosingReport() {
  var historySheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Payment_History");
  if (!historySheet) return null;
  var data = historySheet.getDataRange().getValues();
  var today = new Date();
  var todayStr = Utilities.formatDate(today, Session.getScriptTimeZone(), "dd/MM/yyyy");
  
  var report = { date: todayStr, totalTx: 0, cashTotal: 0, qrTotal: 0, exemptTotal: 0, grandTotal: 0, details: [] };
  
  for (var i = 1; i < data.length; i++) {
    if (!data[i][2]) continue;
    var rowDate = new Date(data[i][2]);
    var rowDateStr = Utilities.formatDate(rowDate, Session.getScriptTimeZone(), "dd/MM/yyyy");
    
    if (rowDateStr === todayStr) {
      var amount = Number(data[i][4]) || 0;
      var method = String(data[i][5]);
      var cashier = String(data[i][6]);
      
      report.totalTx++;
      report.grandTotal += amount;
      
      if (method.indexOf("Cash") !== -1 || method.indexOf("សាច់ប្រាក់") !== -1) {
        report.cashTotal += amount;
      } else if (method.indexOf("KHQR") !== -1 || method.indexOf("ស្កែន") !== -1) {
        report.qrTotal += amount;
      } else {
        report.exemptTotal += amount;
      }
      
      report.details.push({ id: data[i][0], name: data[i][1], phase: data[i][3], amount: amount, method: method, cashier: cashier });
    }
  }
  return report;
}

function getMonthlyClosingReport() {
  var historySheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Payment_History");
  if (!historySheet) return null;
  var data = historySheet.getDataRange().getValues();
  var today = new Date();
  var currentMonth = today.getMonth();
  var currentYear = today.getFullYear();
  var monthNames = ["មករា", "កុម្ភៈ", "មីនា", "មេសា", "ឧសភា", "មិថុនា", "កក្កដា", "សីហា", "កញ្ញា", "តុលា", "វិច្ឆិកា", "ធ្នូ"];
  
  var report = { monthStr: monthNames[currentMonth] + " ឆ្នាំ " + currentYear, totalTx: 0, cashTotal: 0, qrTotal: 0, grandTotal: 0, details: [] };
  
  for (var i = 1; i < data.length; i++) {
    if (!data[i][2]) continue;
    var rowDate = new Date(data[i][2]);
    if (rowDate.getMonth() === currentMonth && rowDate.getFullYear() === currentYear) {
      var amount = Number(data[i][4]) || 0;
      var method = String(data[i][5]);
      report.totalTx++;
      report.grandTotal += amount;
      if (method.indexOf("Cash") !== -1 || method.indexOf("សាច់ប្រាក់") !== -1) {
        report.cashTotal += amount;
      } else if (method.indexOf("KHQR") !== -1 || method.indexOf("ស្កែន") !== -1) {
        report.qrTotal += amount;
      }
      report.details.push({ date: Utilities.formatDate(rowDate, Session.getScriptTimeZone(), "dd/MM/yyyy"), id: data[i][0], name: data[i][1], phase: data[i][3], amount: amount, method: method });
    }
  }
  return report;
}

function generateNextStudentId() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Students_Payment") || ss.getSheets()[0];
  var lastRow = Math.max(sheet.getLastRow(), 2);
  var data = sheet.getRange("A2:A" + lastRow).getValues();
  var maxNumber = 0;
  
  for (var i = 0; i < data.length; i++) {
    var currentId = String(data[i][0]).trim();
    if (currentId.indexOf("AKKNGS-") === 0) {
      var numStr = currentId.replace("AKKNGS-", "");
      var num = parseInt(numStr, 10);
      if (!isNaN(num) && num > maxNumber) maxNumber = num;
    }
  }
  
  var nextNumber = maxNumber + 1;
  var nextNumberStr = nextNumber.toString();
  while (nextNumberStr.length < 6) { nextNumberStr = "0" + nextNumberStr; }
  return "AKKNGS-" + nextNumberStr;
}

function sendTelegramMessage(text) {
  if (TELEGRAM_BOT_TOKEN === "YOUR_BOT_TOKEN_HERE" || TELEGRAM_CHAT_ID === "YOUR_CHAT_ID_HERE") return;
  var url = "https://api.telegram.org/bot" + TELEGRAM_BOT_TOKEN + "/sendMessage";
  var payload = { "chat_id": TELEGRAM_CHAT_ID, "text": text, "parse_mode": "HTML" };
  var options = { "method": "post", "contentType": "application/json", "payload": JSON.stringify(payload), "muteHttpExceptions": true };
  try { UrlFetchApp.fetch(url, options); } catch(e) {}
}
