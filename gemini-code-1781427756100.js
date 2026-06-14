// دالة استقبال طلبات جلب البيانات (GET) من موقع GitHub
function doGet(e) {
  try {
    var action = e.parameter.action;
    var responseData;

    if (action === 'getEvents') {
      responseData = getEventsList();
    } else if (action === 'getEventData') {
      responseData = getEventDataAndStats(e.parameter.eventName);
    } else {
      responseData = { status: "error", message: "Action غير معروف" };
    }

    return ContentService.createTextOutput(JSON.stringify(responseData))
                         .setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({ status: "error", message: error.toString() }))
                         .setMimeType(ContentService.MimeType.JSON);
  }
}

// دالة استقبال طلبات إرسال البيانات (POST) مثل التسجيل وإنشاء الفعاليات
function doPost(e) {
  try {
    var requestData = JSON.parse(e.postData.contents);
    var action = requestData.action;
    var responseData;

    if (action === 'createEvent') {
      responseData = createNewEvent(requestData.university, requestData.date, requestData.lectureName, requestData.templateId);
    } else if (action === 'register') {
      responseData = registerStudent(requestData.eventName, requestData.studentData);
    } else if (action === 'generatePDF') {
      responseData = generateBulkPDF(requestData.eventName);
    } else {
      responseData = { status: "error", message: "Action غير معروف" };
    }

    return ContentService.createTextOutput(JSON.stringify(responseData))
                         .setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({ status: "error", message: error.toString() }))
                         .setMimeType(ContentService.MimeType.JSON);
  }
}

// 1. دالة إنشاء فعالية جديدة
function createNewEvent(university, date, lectureName, templateId) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetName = university + " - " + date;
  
  if (ss.getSheetByName(sheetName)) {
    return { status: "error", message: "هذه الفعالية موجودة بالفعل!" };
  }
  
  var sheet = ss.insertSheet(sheetName);
  var headers = ["م", "الاسم بالكامل", "النوع", "رقم الهاتف", "الجامعة", "الكلية", "الفرقة الدراسية", "التاريخ"];
  sheet.appendRow(headers);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#f3f3f3");
  
  var configSheet = ss.getSheetByName("Events_Config") || ss.insertSheet("Events_Config");
  if (configSheet.getLastRow() === 0) {
    configSheet.appendRow(["اسم الفعالية", "المحاضرة", "رابط الـ Template"]);
  }
  configSheet.appendRow([sheetName, lectureName, templateId]);
  
  return { status: "success", eventName: sheetName };
}

// 2. دالة تسجيل طالب جديد مع منع التكرار
function registerStudent(eventName, studentData) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(eventName);
  
  if (!sheet) {
    return { status: "error", message: "الفعالية غير موجودة." };
  }
  
  var data = sheet.getDataRange().getValues();
  
  // التحقق من تكرار الهاتف
  for (var i = 1; i < data.length; i++) {
    if (data[i][3].toString() === studentData.phone.toString()) {
      return { status: "duplicate", message: "تم تسجيل هذا الرقم مسبقاً في هذه الفعالية." };
    }
  }
  
  var nextId = data.length;
  sheet.appendRow([
    nextId,
    studentData.name,
    studentData.gender,
    "'" + studentData.phone,
    studentData.university,
    studentData.college || "غير محدد",
    studentData.grade,
    studentData.date
  ]);
  
  return { status: "success" };
}

// 3. جلب أسماء الفعاليات
function getEventsList() {
  var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  var list = [];
  for (var i = 0; i < sheets.length; i++) {
    var name = sheets[i].getName();
    if (name !== "Events_Config") list.push(name);
  }
  return list;
}

// 4. جلب بيانات وإحصائيات الفعالية
function getEventDataAndStats(eventName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(eventName);
  if (!sheet) return { status: "error", message: "الفعالية غير موجودة" };
  
  var data = sheet.getDataRange().getValues();
  var rows = data.slice(1);
  
  var stats = { total: rows.length, male: 0, female: 0 };
  
  rows.forEach(function(row) {
    if (row[2] === "طالب") stats.male++;
    if (row[2] === "طالبة") stats.female++;
  });
  
  return { students: rows, stats: stats };
}

// 5. دمج الشهادات وتوليد ملف PDF مجمع
function generateBulkPDF(eventName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var configSheet = ss.getSheetByName("Events_Config");
  var templateId = "";
  
  if (configSheet) {
    var configData = configSheet.getDataRange().getValues();
    for (var i = 1; i < configData.length; i++) {
      if (configData[i][0] === eventName) {
        templateId = configData[i][2];
        break;
      }
    }
  }
  
  if (!templateId) return { status: "error", message: "لم يتم العثور على قالب الشهادة لهذه الفعالية." };
  
  var studentSheet = ss.getSheetByName(eventName);
  var students = studentSheet.getDataRange().getValues().slice(1);
  
  if (students.length === 0) return { status: "error", message: "لا يوجد حضور لتوليد شهادات لهم." };
  
  var templateFile = DriveApp.getFileById(templateId);
  var tempPresentationFile = templateFile.makeCopy("Temp_" + eventName);
  var presentation = SlidesApp.openById(tempPresentationFile.getId());
  var masterSlide = presentation.getSlides()[0];
  
  students.forEach(function(student) {
    var newSlide = masterSlide.duplicate();
    newSlide.replaceAllText("{{الاسم}}", student[1]);
    newSlide.replaceAllText("{{التاريخ}}", student[7]);
  });
  
  masterSlide.remove();
  presentation.saveAndClose();
  
  var pdfBlob = tempPresentationFile.getAs(MimeType.PDF);
  pdfBlob.setName("Certificates_" + eventName.replace(/ /g, "_") + ".pdf");
  var pdfFile = DriveApp.createFile(pdfBlob);
  
  DriveApp.getFileById(tempPresentationFile.getId()).setTrashed(true);
  
  return { status: "success", downloadUrl: pdfFile.getDownloadUrl() };
}