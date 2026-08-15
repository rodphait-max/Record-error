function doGet() {
  return HtmlService.createHtmlOutput('API Service is Running');
}

function doPost(e) {
  try {
    var payload = JSON.parse(e.postData.contents);
    var result = saveData(payload);
    
    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ 
      status: "error", 
      message: err.message 
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

function saveData(payload) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var formType = payload.formType;
    var data = payload.data;
    var sheetName = "";

    if (formType === 'operationError') {
      sheetName = "บันทึกข้อผิดพลาดจากการปฏิบัติงาน";
      var sheet = getOrCreateSheet(ss, sheetName, [
        "วันที่", "คนจัดสินค้า", "คนสแกนสินค้า", "คนแพ็ค/โหลด", "Delivery", 
        "ชื่อลูกค้า", "ขนส่ง", "Article", "สาเหตุ", "รายละเอียดเพิ่มเติม", "ผู้บันทึก"
      ]);
      
      sheet.appendRow([
        data.logDate || '',
        data.picker || '',
        data.scanner || '',
        data.packerLoader || '',
        data.delivery || '',
        data.customerName || '',
        data.carrier || '',
        data.article || '',
        data.cause || '',
        data.details || '',
        data.recordedBy || ''
      ]);

    } else if (formType === 'putaway') {
      sheetName = "บันทึกการเก็บสินค้า";
      var sheet = getOrCreateSheet(ss, sheetName, [
        "วันที่", "คนแจ้งปัญหา", "Article", "ปัญหา", "Bin ที่พบ", "วิธีแก้ปัญหา", "รายละเอียดเพิ่มเติม", "พนักงานที่แก้ไข", "คนบันทึก"
      ]);

      sheet.appendRow([
        data.logDate || '',
        data.reporter || '',
        data.article || '',
        data.problem || '',
        data.binFound || '',
        data.solution || '',
        data.details || '',
        data.resolvedBy || '',
        data.recordedBy || ''
      ]);
    }

    return { status: "success", sheet: sheetName };
  } catch (err) {
    return { status: "error", message: err.message };
  }
}

function getOrCreateSheet(ss, sheetName, headers) {
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#f1f5f9");
  }
  return sheet;
}