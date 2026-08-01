// ==========================================
// 1. API TRA CỨU ĐƠN HÀNG (doGet) - TÌM BẰNG SĐT
// ==========================================
function doGet(e) {
  var sdt = e.parameter.sdt || e.parameter.phone || e.parameter.maDon; 
  if (!sdt) return ContentService.createTextOutput(JSON.stringify({status: "error", message: "Vui lòng nhập số điện thoại hoặc mã đơn"})).setMimeType(ContentService.MimeType.JSON);
  
  sdt = sdt.toString().replace(/\D/g, '');
  
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  var data = sheet.getDataRange().getDisplayValues();

  for (var i = data.length - 1; i > 0; i--) {
    var currentPhone = data[i][2].toString().replace(/\D/g, ''); 
    
    if (currentPhone === sdt && currentPhone !== "") {
      var name = data[i][1];      
      var phone = data[i][2].replace(/'/g, ""); 
      var address = data[i][6] + ", " + data[i][5] + ", " + data[i][4] + ", " + data[i][3];
      var balance = data[i][11]; // L(11)
      var note = data[i][12];    // M(12)
      var maVanDon = data[i][13];// N(13)
      var status = data[i][15];  // P(15)
      
      var isPaid = false;
      if (status && (status.toString().toUpperCase().indexOf("ĐÃ THANH TOÁN") !== -1 || status.toString().indexOf("✅") !== -1)) {
         isPaid = true;
         balance = "0";
      }

      return ContentService.createTextOutput(JSON.stringify({
          status: "success", 
          maVanDon: maVanDon || "", 
          name: name, 
          phone: phone, 
          address: address, 
          balance: balance, 
          note: note,
          isPaid: isPaid
      })).setMimeType(ContentService.MimeType.JSON);
    }
  }
  return ContentService.createTextOutput(JSON.stringify({status: "not_found", message: "Không tìm thấy đơn hàng với số điện thoại này!"})).setMimeType(ContentService.MimeType.JSON);
}

// ==========================================
// 2. BỘ ĐỊNH TUYẾN (doPost)
// ==========================================
function doPost(e) {
  try {
    if (e.postData && e.postData.type === "application/json") {
      var contents = e.postData.contents;
      var data = JSON.parse(contents);
      if (data.content || data.description || data.transferAmount) {
        return handleSePayWebhook(data);
      }
    }
    if (e.parameter && e.parameter.name) {
      return handleOrderForm(e);
    }
    return ContentService.createTextOutput(JSON.stringify({ "success": false, "error": "No valid data found" })).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ "success": false, "error": err.toString() })).setMimeType(ContentService.MimeType.JSON);
  }
}

// ==========================================
// 3. HÀM XỬ LÝ ĐƠN HÀNG MỚI (Từ Web)
// ==========================================
function handleOrderForm(e) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  const p = e.parameter;
  
  const totalValue = parseFloat(p.total || 0);
  const paidValue = parseFloat(p.paid || 0);
  const balanceValue = totalValue - paidValue;

  // LƯU Ý LOGIC Ở ĐÂY: Nếu chọn CK toàn bộ (balanceValue = 0) thì hệ thống tự lưu là Đã thanh toán.
  const trangThai = (balanceValue <= 0) ? "✅ ĐÃ THANH TOÁN (Toàn bộ)" : "Chưa thanh toán (Mới cọc)";

  sheet.appendRow([
    new Date(),        
    p.name,            
    "'" + p.phone,     
    p.province,        
    p.district,        
    p.ward,            
    p.address,         
    p.source,          
    p.account,         
    totalValue,        
    paidValue,         
    balanceValue,      
    p.note,            
    "",                
    "",                
    trangThai          
  ]);
  return ContentService.createTextOutput("Success");
}

// ==========================================
// 4. HÀM TỰ ĐỘNG GẠCH NỢ TỪ SEPAY - TÌM THEO SĐT
// ==========================================
function handleSePayWebhook(data) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  const dataRange = sheet.getDataRange();
  const values = dataRange.getValues();
  
  const noiDungCK = (data.content || data.description || "").toUpperCase();
  const soTienCK = data.transferAmount || data.amount || 0;
  
  for (let i = values.length - 1; i > 0; i--) {
    let phoneTrongSheet = values[i][2].toString().replace(/\D/g, ''); 
    
    if (phoneTrongSheet === "") continue;
    
    if (noiDungCK.indexOf(phoneTrongSheet) !== -1) {
      sheet.getRange(i + 1, 16).setValue("✅ ĐÃ THANH TOÁN: " + soTienCK.toLocaleString() + "đ");
      SpreadsheetApp.flush();
      return ContentService.createTextOutput(JSON.stringify({success: true})).setMimeType(ContentService.MimeType.JSON);
    }
  }
  return ContentService.createTextOutput(JSON.stringify({success: false, message: "Không tìm thấy số điện thoại trong nội dung CK"}));
}
