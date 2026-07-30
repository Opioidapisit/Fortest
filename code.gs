// @ts-nocheck
// ***************************************************************************************************
// *** ส่วนที่ 1: WEB APP CONFIG & API (ส่วนใหม่สำหรับหน้าเว็บ) ***
// ***************************************************************************************************

function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
      .setTitle('ระบบสั่งซื้อสินค้า NUVITRA')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no');
}

// ฟังก์ชันค้นหาลูกค้าเก่า
function searchCustomer(phone) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("การตอบแบบฟอร์ม 1"); // 🚨 ตรวจสอบชื่อ Sheet ว่าตรงไหม
  var data = sheet.getDataRange().getValues();
  
  // วนลูปจากล่างขึ้นบน
  for (var i = data.length - 1; i >= 1; i--) {
    // เทียบเบอร์โทร (ตัดขีดตัดวรรคออก)
    if (String(data[i][3]).replace(/[- ]/g,'') == String(phone).replace(/[- ]/g,'')) {
      return {
        found: true,
        name: data[i][1],       // Col B
        address: data[i][2],    // Col C
        email: data[i][25],     // Col Y (New Index)
        taxName: data[i][27],   // Col AA (New Index)
        taxId: data[i][28],     // Col AB (New Index)
        taxAddr: data[i][29]    // Col AC (New Index)
      };
    }
  }
  return { found: false };
}

// ฟังก์ชันรับค่าจากหน้าเว็บ -> บันทึก -> สั่งออกบิล
function submitWebForm(formObject) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) throw new Error("ระบบกำลังทำงาน โปรดรอสักครู่");

  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("การตอบแบบฟอร์ม 1"); // 🚨 ตรวจสอบชื่อ Sheet
    
    // วิธีหา LastRow แบบใหม่ (ป้องกัน Ghost Rows)
    var allData = sheet.getRange("A1:A").getValues();
    var lastRow = 0;
    for (var i = allData.length - 1; i >= 0; i--) {
      if (allData[i][0] !== "" && allData[i][0] !== null) {
        lastRow = i + 1;
        break;
      }
    }
    if (lastRow === 0) lastRow = 1;

    var newRow = lastRow + 1;
    var lastCol = sheet.getLastColumn(); // ใช้ LastColumn จาก Sheet เพื่อความชัวร์ (ควรจะไปถึง Col CE หรือมากกว่า)
    
    // 1. ก๊อปปี้สูตรจากแถวบนลงมา (สำคัญมาก ห้ามลืม!)
    sheet.getRange(lastRow, 1, 1, lastCol).copyTo(sheet.getRange(newRow, 1));
    
    // 2. เตรียมข้อมูลลงบันทึก
    var timestamp = new Date();
    // สร้าง Array ให้ยาวพอที่จะคลุมข้อมูล Input ทั้งหมด (A ถึง AD = 29 คอลัมน์)
    var values = new Array(30); 

    // A-D (ข้อมูลพื้นฐาน)
    values[0] = timestamp;
    values[1] = "'" + formObject.name;
    values[2] = formObject.address;
    values[3] = "'" + formObject.phone;
    
    // E-X (สินค้า 20 ตัว)
    values[4] = formObject.p1 || "";
    values[5] = formObject.p2 || "";
    values[6] = formObject.p3 || "";
    values[7] = formObject.p4 || "";
    values[8] = formObject.p5 || "";
    values[9] = formObject.p6 || "";
    values[10] = formObject.p7 || "";
    values[11] = formObject.p8 || "";
    values[12] = formObject.p9 || "";
    values[13] = formObject.p10 || "";
    values[14] = formObject.p11 || "";
    values[15] = formObject.p12 || "";
    values[16] = formObject.p13 || "";
    values[17] = formObject.p14 || "";
    values[18] = formObject.p15 || "";
    values[19] = formObject.p16 || "";
    values[20] = formObject.p17 || "";
    values[21] = formObject.p18 || "";
    values[22] =  "";
    values[23] = formObject.p19 || "";
    values[24] = formObject.p20 || "";
    
    // Y-AC (ข้อมูลเสริม & ภาษี)
    values[25] = formObject.email || "";
    values[26] = formObject.taxRequest;
    if (formObject.taxRequest === "ต้องการ") {
        values[27] = formObject.taxName;
        values[28] = "'" + formObject.taxId;
        values[29] = formObject.taxAddr;
    } else {
        values[27] = ""; values[28] = ""; values[29] = "";
    }
    
    // 3. บันทึกข้อมูลเฉพาะส่วนที่เป็น Input ลงไปทับสูตร (Col A ถึง AD)
    // เขียนข้อมูลแถวเดียว (1 row) จำนวน 29 คอลัมน์ (1-30)
    sheet.getRange(newRow, 1, 1, 30).setValues([values]);

    // 👇 วางโค้ดใหม่ตรงนี้
    sheet.getRange(newRow, 87, 1, 2).setValues([[
    formObject.p21 || "",
    formObject.p22 || ""
    ]]);


    SpreadsheetApp.flush(); // บังคับคำนวณเดี๋ยวนี้
    Utilities.sleep(3000); // รอสูตรคำนวณ 3 วินาที

    // 4. สั่ง Process งานเดิม (ออกบิล/ส่งเมล)
    processOrder(sheet, newRow);
    
    return { success: true };

  } catch (error) {
    return { success: false, error: error.message };
  } finally {
    lock.releaseLock();
  }
}

// ***************************************************************************************************
// *** ส่วนที่ 2: BUSINESS LOGIC (Mapping ใหม่ 20 สินค้า) ***
// ***************************************************************************************************

function processOrder(sheet, row) { 
  
  // Configuration
  var DOC_TEMPLATE_ID = '1QumBo0I9e5wKL39Wy8sNFWO5O1Qg9hfZyx42N08cfuk';  // Template ใหม่
  var TAX_DOC_TEMPLATE_ID = '1bIBUXqoOEhjR8RiXX6y5gRcIz9wDMUp0TCgARRa3yck'; 
  var yourEmail = "cmo.trang@gmail.com";
  var LINE_ACCESS_TOKEN = "qkiFib2YlSQYqxEU9AnjWCHu1s515VTw9zq6/rulUjewmltBmdcG6WLFcPn+e2lTCvHhlU/0arZcJilOpP1PC81N+wFiv3qjvcSMZtY00fTMkerOqUVavRHObMaP2acwfXdTRaTtMz8wLxpHa2W36QdB04t89/1O/w1cDnyilFU="; 
  var LINE_USER_ID = "C90ea85531dc7ffa9079a7fcbc0c7f4fc"
  
  // --- MAPPING INDICES (Updated based on 20 products) ---
  // A=0, B=1, ... Y=24, Z=25, AA=26, AB=27, A  C=28
  // Qty: E(4) - X(23)
  // Price Total per Item: AG(32) - AZ(51)  (AG is col 33 -> index 32)
  // Unit Price (PU): BL(63) - CE(82) (BL is col 64 -> index 63)
  
  // Indices Array for Doc Replacement (Must match templateTags order)
  var dataIndices = [
  0, 1, 2, 3,
  4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21,
  23, 24, 86, 87,
  28, 27, 29, 31,
  33, 34, 35, 36, 37, 38, 39, 40, 41, 42, 43, 44, 45, 46, 47, 48, 49, 50,
  52, 53, 88, 89,
  54,
  55, 57, 58, 61, 64, 63, 59, 60, 62,
  65, 66, 67, 68, 69, 70, 71, 72, 73, 74, 75, 76, 77, 78, 79, 80, 81, 82,
  84, 85, 90, 91
];

 var templateTags = [
  "Date", "customer_name", "customer_shipping", "customer_mobile",
  "product_1", "product_2", "product_3", "product_4", "product_5", "product_6", "product_7", "product_8", "product_9", "product_10", "product_11", "product_12", "product_13", "product_14", "product_15", "product_16", "product_17", "product_18",
  "product_19", "product_20", "product_21", "product_22",
  "TaxID", "TaxName", "TaxAddress", "product_sum",
  "price_1", "price_2", "price_3", "price_4", "price_5", "price_6", "price_7", "price_8", "price_9", "price_10", "price_11", "price_12", "price_13", "price_14", "price_15", "price_16", "price_17", "price_18",
  "price_19", "price_20", "price_21", "price_22",
  "price_sum",
  "price_ship", "VAT", "NOVAT", "Asta", "7VAT", "PriceVAT", "DC", "price_total", "BathTEXT",
  "PU_1", "PU_2", "PU_3", "PU_4", "PU_5", "PU_6", "PU_7", "PU_8", "PU_9", "PU_10", "PU_11", "PU_12", "PU_13", "PU_14", "PU_15", "PU_16", "PU_17", "PU_18",
  "PU_19", "PU_20", "PU_21", "PU_22"
];

  // Product Map for Tax Invoice Table
 var productMap = [
  { tag: "product_1", name: "Nuvitra Acerola Cherry Plus", puTag: "PU_1" },
  { tag: "product_2", name: "Nuvitra Calcium L-threonate", puTag: "PU_2" },
  { tag: "product_3", name: "Nuvitra Collagen Plus", puTag: "PU_3" },
  { tag: "product_4", name: "Nuvitra Fish Oil 1200 mg", puTag: "PU_4" },
  { tag: "product_5", name: "Nuvitra Harigro", puTag: "PU_5" },
  { tag: "product_6", name: "Nuvitra L-Glutathione Plus", puTag: "PU_6" },
  { tag: "product_7", name: "Nuvitra Lutein & Zeaxanthin", puTag: "PU_7" },
  { tag: "product_8", name: "Nuvitra Magnesium Plus", puTag: "PU_8" },
  { tag: "product_9", name: "Nuvitra Multivitamin Plus Minerals", puTag: "PU_9" },
  { tag: "product_10", name: "Nuvitra Astaxanthin Plus", puTag: "PU_10" },
  { tag: "product_11", name: "Nuvitra Zinc Plus C", puTag: "PU_11" },
  { tag: "product_12", name: "Nuvitra Date Palm Plus", puTag: "PU_12" },
  { tag: "product_13", name: "Nuvitra Colla Cal plus", puTag: "PU_13" },
  { tag: "product_14", name: "Nuvitra Gotu Kola plus", puTag: "PU_14" },
  { tag: "product_15", name: "Nuvitra Vitamin C 1000 mg plus", puTag: "PU_15" },
  { tag: "product_16", name: "Nuvitra B complex plus minerals", puTag: "PU_16" },
  { tag: "product_17", name: "Nuvitra Calcium plus (Box)", puTag: "PU_17" },
  { tag: "product_18", name: "Nuvitra Magnesium plus (Box)", puTag: "PU_18" },
  { tag: "product_19", name: "Nuvitra Multivitamins plus minerals (Box)", puTag: "PU_19" },
  { tag: "product_20", name: "Nuvitra Lutein&Zeaxanthin plus (Box)", puTag: "PU_20" },
  { tag: "product_21", name: "Nuvitra Q10 plus", puTag: "PU_21" },
  { tag: "product_22", name: "Nuvitra L-Arginine 1000 mg plus", puTag: "PU_22" }
];

  // ดึงข้อมูลจากแถว
  var data = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];
  
  // ดึงตัวแปรสำคัญ
  var customerEmail = data[25]; // Col Y
  var taxId = data[28]; // Col AB
  var customerName = data[1]; // Col B
  var customerShipping = data[2]; // Col C
  var customerMobile = data[3]; // Col D

  // 1. สร้างใบส่งของ (PDF)
  var templateFile = DriveApp.getFileById(DOC_TEMPLATE_ID);
  var tempFile = templateFile.makeCopy(); 
  var tempDoc = DocumentApp.openById(tempFile.getId());
  var body = tempDoc.getBody();

  for (var i = 0; i < dataIndices.length; i++) {
    var index = dataIndices[i];
    var tag = "{{" + templateTags[i] + "}}";
    var value = data[index];
    if (value === undefined || value === null) value = "";
    if (templateTags[i] === "Date" && value) value = new Date(value).toLocaleDateString('th-TH');
    
    // Formatting Numbers for Price tags
    if (typeof templateTags[i] === 'string' && (templateTags[i].startsWith("price_") || templateTags[i].includes("VAT") || templateTags[i] === "Discount" || templateTags[i] === "BathTEXT")) {
          // 🛑 FIX: ยกเว้นคำว่า VAT และ NOVAT ออกจากการจัดรูปแบบตัวเลข เพราะมันเป็นข้อความ
          if(templateTags[i] !== "BathTEXT" && templateTags[i] !== "VAT" && templateTags[i] !== "NOVAT") {
             value = formatCurrency(parseFloat(value) || 0);
          }
    }
    
    body.replaceText(tag, value);
  }
  
  tempDoc.saveAndClose();
  
  // แปลงเป็น PDF
  var tempDocFile = DriveApp.getFileById(tempFile.getId());
  var pdfName = "Invoice_" + (customerName ? customerName : "NoName") + ".pdf";
  var pdfBlob = tempDocFile.getAs('application/pdf');
  pdfBlob.setName(pdfName);

  var actualPdfFile = DriveApp.createFile(pdfBlob); 
  actualPdfFile.setSharing(DriveApp.Access.ANYONE, DriveApp.Permission.VIEW);
  var pdfUrl = actualPdfFile.getUrl(); 
  tempFile.setTrashed(true);

  // 2. สร้างใบกำกับภาษี (ถ้ามี)
  var taxDocUrl = null;
  if (taxId && TAX_DOC_TEMPLATE_ID !== 'YOUR_TAX_INVOICE_TEMPLATE_ID') { 
    
    var purchasedItems = [];
    var productTagStart = templateTags.indexOf("product_1"); 
    
    for (var i = 0; i < productMap.length; i++) {
      var itemMap = productMap[i];
      var productDataIndex = dataIndices[productTagStart + i];
      var quantity = parseFloat(data[productDataIndex]) || 0;
      
      if (quantity > 0) {
        var puTagIndex = templateTags.indexOf(itemMap.puTag);
        var unitPriceIndex = dataIndices[puTagIndex];
        var unitPrice = parseFloat(data[unitPriceIndex]) || 0;
        
        purchasedItems.push({
          name: itemMap.name,
          quantity: quantity,
          unitPrice: unitPrice, 
          totalPrice: quantity * unitPrice
        });
      }
    }
// เพิ่มค่าจัดส่งจาก Column BB (Index 53) ไปต่อท้ายรายการสินค้า
    var shippingCost = parseFloat(data[55]) || 0; 
    if (shippingCost > 0) {
      purchasedItems.push({
        name: "ค่าจัดส่ง",
        quantity: 1,
        unitPrice: shippingCost,
        totalPrice: shippingCost
      });
    }
    
    var taxTemplateFile = DriveApp.getFileById(TAX_DOC_TEMPLATE_ID);
    var tempTaxFile = taxTemplateFile.makeCopy();
    var tempTaxDoc = DocumentApp.openById(tempTaxFile.getId());
    var taxBody = tempTaxDoc.getBody();

    // แทนที่ Tags
    for (var i = 0; i < dataIndices.length; i++) {
        var index = dataIndices[i];
        var tag = "{{" + templateTags[i] + "}}";
        var value = data[index];
        if (value === undefined || value === null) value = "";
        if (templateTags[i] === "Date" && value) value = new Date(value).toLocaleDateString('th-TH');
        else if (typeof templateTags[i] === 'string' && (templateTags[i].startsWith("price_") || templateTags[i].includes("VAT") || templateTags[i].startsWith("PU_") || templateTags[i] === "Discount")) {
             // 🛑 FIX: เพิ่มการยกเว้นในส่วนใบกำกับภาษีด้วย เพื่อความปลอดภัย (เผื่ออนาคตใช้ Tag เหล่านี้)
             if(templateTags[i] !== "VAT" && templateTags[i] !== "NOVAT") {
                 value = formatCurrency(parseFloat(value) || 0);
             }
        }
        taxBody.replaceText(tag, value);
    }

    // สร้างตารางสินค้าใน Word
    var placeholderText = "{{PRODUCTS_TABLE_PLACEHOLDER}}";
    var placeholder = taxBody.findText(placeholderText);
    if (placeholder) {
      var newTable = [["ลำดับ", "รายการสินค้า", "จำนวน (หน่วย)", "ราคาต่อหน่วย", "จำนวนเงิน"]];
      for (var j = 0; j < purchasedItems.length; j++) {
        var item = purchasedItems[j];
        newTable.push([
          (j + 1).toString(), 
          item.name,          
          item.quantity.toString() + (item.name === "ค่าจัดส่ง" ? " (หน่วย)" : " (ชิ้น)"), 
          formatCurrency(item.unitPrice), 
          formatCurrency(item.totalPrice) 
        ]);
      }
      var insertionElement = placeholder.getElement().getParent();
      var tableElement = insertionElement.getParent().insertTable(insertionElement.getParent().getChildIndex(insertionElement), newTable);
      insertionElement.removeFromParent();
      
      // จัด Format ตาราง
      var headerRow = tableElement.getRow(0);
      tableElement.setColumnWidth(0, 50);
      tableElement.setColumnWidth(1, 220);
      tableElement.setColumnWidth(2, 70);
      tableElement.setColumnWidth(3, 80);
      tableElement.setColumnWidth(4, 80);
      for (var k = 0; k < headerRow.getNumCells(); k++) {
        headerRow.getCell(k).setText(headerRow.getCell(k).getText()).setBold(true);
      }
    }
    tempTaxDoc.saveAndClose(); 
    var taxDocName = "TAX_Invoice_" + (customerName ? customerName : "NoName") + ".docx";
    var taxDocFile = DriveApp.getFileById(tempTaxFile.getId());
    taxDocFile.setName(taxDocName);
    taxDocFile.setSharing(DriveApp.Access.ANYONE, DriveApp.Permission.VIEW);
    taxDocUrl = taxDocFile.getUrl();
  }
  
  // 3. ส่ง Email (Admin & Customer)
  var subjectAdmin = "มีคำสั่งซื้อใหม่ จาก: " + customerName;
  var bodyMessageAdmin = "มีคำสั่งซื้อใหม่ โปรดตรวจสอบใบส่งสินค้า:\n\n" + pdfUrl + "\n\nขอบคุณครับ";
  
  MailApp.sendEmail({
    to: yourEmail,
    subject: subjectAdmin,
    body: bodyMessageAdmin,
    attachments: [pdfBlob]
  });

  if (customerEmail && customerEmail.indexOf('@') > -1) { 
    var subjectCustomer = "ใบสรุปคำสั่งซื้อ Nuvitra ของคุณ";
    var bodyMessageCustomer = "เรียนคุณ " + customerName + ",\n\nขอบคุณสำหรับคำสั่งซื้อ โปรดคลิกลิงก์ด้านล่างเพื่อดูใบเสร็จรับเงินฉบับเต็ม:\n\n" + pdfUrl + "\n\nขอบคุณครับ"+ "\n\nหากพบปัญหา สามารถติดต่อบั๊มได้โดยตรงผ่านไลน์นี้ครับ https://line.me/ti/p/l1pnYMW0PR";;
    MailApp.sendEmail({
      to: customerEmail, 
      subject: subjectCustomer,
      body: bodyMessageCustomer,
      attachments: [pdfBlob] 
    });
  }
  
  // 4. แจ้งเตือน LINE
  var lineMessage = "\n🔔คำสั่งซื้อ Nuvitra ใหม่🔔"
                  + "\nชื่อลูกค้า : " + (customerName ? customerName : "ไม่ระบุชื่อ")
                  + "\nจัดส่งที่ : " + (customerShipping ? customerShipping : "ไม่ระบุที่อยู่")
                  + "\nโทรศัพท์ : " + (customerMobile ? customerMobile : "ไม่ระบุเบอร์โทร");
  
  if (taxDocUrl) {
    lineMessage += "\n--------------------------"
                 + "\n⭐ มีการร้องขอใบกำกับภาษี ⭐"
                 + "\nคลิกดูไฟล์ DOC : " + taxDocUrl;
  }
  
  lineMessage += "\n--------------------------"
               + "\nคลิกดูใบส่งของ (PDF) : " + pdfUrl;
  
  sendLineMessage(lineMessage, LINE_USER_ID, LINE_ACCESS_TOKEN);
}

// ฟังก์ชัน Helper
function formatCurrency(n) {
  if (typeof n !== 'number' || isNaN(n) || n === null || n === "") return '0.00';
  return n.toFixed(2).replace(/\B(?=(\d{3})+(?!\d))/g, ",");
}

function sendLineMessage(message, userId, accessToken) {
  var lineUrl = "https://api.line.me/v2/bot/message/push";
  var options = {
    "method": "post",
    "headers": {
      "Content-Type": "application/json",
      "Authorization": "Bearer " + accessToken
    },
    "payload": JSON.stringify({
      "to": userId,
      "messages": [{ "type": "text", "text": message }]
    }),
    "muteHttpExceptions": true // สำคัญ: ให้ดักจับ error response ได้เอง แทนที่จะ throw ทันที
  };

  var maxRetries = 3;
  var attempt = 0;

  while (attempt < maxRetries) {
    attempt++;
    try {
      var response = UrlFetchApp.fetch(lineUrl, options);
      var statusCode = response.getResponseCode();

      if (statusCode === 200) {
        // สำเร็จ ไม่ต้อง retry
        return true;
      }

      // ถ้าเป็น error ฝั่งเซิร์ฟเวอร์ LINE (5xx) ให้ลองใหม่
      if (statusCode >= 500 && statusCode < 600) {
        Logger.log("LINE API ตอบกลับ " + statusCode + " (ครั้งที่ " + attempt + "/" + maxRetries + "): " + response.getContentText());
        if (attempt < maxRetries) {
          // หน่วงเวลาแบบ exponential backoff: 2s, 4s, 8s
          Utilities.sleep(Math.pow(2, attempt) * 1000);
          continue;
        }
      } else {
        // Error แบบอื่น (เช่น 400, 401) ไม่ใช่ปัญหาชั่วคราว ไม่ต้อง retry
        Logger.log("LINE API Error (รหัส " + statusCode + "): " + response.getContentText());
        return false;
      }

    } catch (e) {
      // ข้อผิดพลาดระดับเครือข่าย (เช่น timeout)
      Logger.log("LINE Messaging API Exception (ครั้งที่ " + attempt + "/" + maxRetries + "): " + e.message);
      if (attempt < maxRetries) {
        Utilities.sleep(Math.pow(2, attempt) * 1000);
        continue;
      }
    }
  }

  // ลองครบทุกครั้งแล้วยังไม่สำเร็จ
  Logger.log("LINE Messaging API: ส่งข้อความล้มเหลวหลังจากลอง " + maxRetries + " ครั้ง");
  return false;
}
