function editMultipleSheetsAndModifyContent() {
    var folderId = '1LsLN9HoQEOyfDlYF2MXTzR6dzZheOguM'; // ใส่ไอดีของโฟลเดอร์ที่มีไฟล์ Google Sheets
    var folder = DriveApp.getFolderById(folderId);
  
    // สร้างอาร์เรย์ของไฟล์ในโฟลเดอร์
    var filesArray = [];
    var files = folder.getFilesByType(MimeType.GOOGLE_SHEETS);
  
    while (files.hasNext()) {
      filesArray.push(files.next());
    }
  
    // เรียงลำดับไฟล์ตามชื่อไฟล์
    filesArray.sort(function(file1, file2) {
      return file1.getName().localeCompare(file2.getName());
    });
  
    var districtCode = 100; // เริ่มต้นรหัสอำเภอที่ไฟล์แรกเป็น 100
  
    // ลูปผ่านไฟล์ที่เรียงลำดับแล้ว
    for (var i = 0; i < filesArray.length; i++) {
      var file = filesArray[i];
      var sheet = SpreadsheetApp.open(file).getActiveSheet();
  
      // ลบแถวที่ 1, 2, 3 และแถวที่ 5
      if (sheet.getLastRow() >= 5) {
        sheet.deleteRow(5);
      }
      if (sheet.getLastRow() >= 3) {
        sheet.deleteRows(1, 3);
      }
  
      // ลบคอลัมน์ D (คอลัมน์ที่ 4)
      sheet.deleteColumn(4);
  
      // เพิ่มคอลัมน์ใหม่
      sheet.insertColumnAfter(3); // เพิ่มคอลัมน์ D
      sheet.getRange("D1").setValue("เดือน");
      sheet.insertColumnAfter(4); // เพิ่มคอลัมน์ E
      sheet.getRange("E1").setValue("ปี");
      sheet.insertColumnAfter(5); // เพิ่มคอลัมน์ F
      sheet.getRange("F1").setValue("รหัสอำเภอ");
  
      // เพิ่มค่าในคอลัมน์ D, E, F, G สำหรับทุกแถวที่มีข้อมูล
      var lastRow = sheet.getLastRow();
      for (var j = 2; j <= lastRow; j++) {
        var currentValueD = sheet.getRange(j, 4).getValue(); // เดือน
        sheet.getRange(j, 4).setValue(currentValueD + 1); // เพิ่มค่าเดือน
  
        var currentValueE = sheet.getRange(j, 5).getValue(); // ปี
        sheet.getRange(j, 5).setValue(currentValueE + 2564); // เพิ่มค่าปี
  
        // ตั้งรหัสอำเภอในทุกแถวให้เป็นค่าเดียวกัน (districtCode)
        sheet.getRange(j, 6).setValue(districtCode); // รหัสอำเภอ
      }
  
      districtCode++; // เพิ่มรหัสอำเภอสำหรับไฟล์ถัดไป
    }
  }