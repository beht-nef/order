// ==========================================
// 1. API TRA CỨU ĐƠN HÀNG (doGet) - TÌM BẰNG SĐT
// ==========================================
function doGet(e) {
  // Đổi từ maDon sang sdt hoặc phone
  var sdt = e.parameter.sdt || e.parameter.phone; 
  if (!sdt) return ContentService.createTextOutput(JSON.stringify({status: "error", message: "Vui lòng nhập số điện thoại"})).setMimeType(ContentService.MimeType.JSON);

  // Chỉ lấy các chữ số trong chuỗi truyền vào để so sánh chính xác
  sdt = sdt.toString().replace(/\D/g, ''); 
  
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  var data = sheet.getDataRange().getDisplayValues();

  // Chạy vòng lặp từ dưới lên (lấy đơn mới nhất nếu có khách đặt nhiều lần trùng 1 SĐT)
  for (var i = data.length - 1; i > 0; i--) {
    var currentPhone = data[i][3].toString().replace(/\D/g, ''); // Cột D (SĐT)
    
    if (currentPhone === sdt && currentPhone !== "") {
      var maVanDon = data[i][15]; // Cột P
      var name = data[i][2]; 
      var phone = data[i][3].replace(/'/g, ""); 
      var address = data[i][7] + ", " + data[i][6] + ", " + data[i][5] + ", " + data[i][4]; 
      var balance = data[i][12]; // Cột M: Số tiền còn lại
      var payment = data[i][13]; // Cột N: Hình thức thanh toán
      var status = data[i][17];  // Cột R: TRẠNG THÁI THANH TOÁN
      
      var isPaid = false;
      if (status && (status.toString().toUpperCase().indexOf("ĐÃ THANH TOÁN") !== -1 || status.toString().indexOf("✅") !== -1)) {
         isPaid = true;
         balance = "0 VNĐ (Đã thanh toán)";
      }

      return ContentService.createTextOutput(JSON.stringify({
          status: "success", 
          maVanDon: maVanDon || "", 
          name: name, 
          phone: phone, 
          address: address, 
          balance: balance, 
          payment: payment,
          isPaid: isPaid
      })).setMimeType(ContentService.MimeType.JSON);
    }
  }
  return ContentService.createTextOutput(JSON.stringify({status: "not_found", message: "Không tìm thấy đơn hàng với số điện thoại này!"})).setMimeType(ContentService.MimeType.JSON);
}

// ==========================================
// 2. BỘ ĐỊNH TUYẾN (doPost) - GIỮ NGUYÊN
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

    return ContentService.createTextOutput(JSON.stringify({ "success": false, "error": "No valid data found" }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ "success": false, "error": err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ==========================================
// 3. HÀM XỬ LÝ ĐƠN HÀNG MỚI (Từ Web) - ĐÃ XÓA ORDER ID
// ==========================================
function handleOrderForm(e) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  const p = e.parameter;
  
  const totalValue = parseFloat(p.total || 0);
  const paidValue = parseFloat(p.paid || 0);
  const balanceValue = totalValue - paidValue;

  sheet.appendRow([
    new Date(),        // A: Thời gian
    "",                // B: ĐỂ TRỐNG (Đã xóa Order ID nhưng giữ khoảng trống để không lệch cột)
    p.name,            // C: Tên người nhận
    "'" + p.phone,     // D: SĐT (Thêm ' để không bị mất số 0)
    p.province,        // E: Tỉnh
    p.district,        // F: Huyện
    p.ward,            // G: Xã
    p.address,         // H: Địa chỉ chi tiết
    p.source,          // I: Inbox qua đâu
    p.account,         // J: Tên nick FB/Zalo
    totalValue,        // K: Tổng tiền
    paidValue,         // L: Đã cọc
    balanceValue,      // M: Còn lại
    p.payment,         // N: Hình thức TT
    p.note,            // O: Ghi chú
    "",                // P: Mã vận đơn (trống)
    "",                // Q: ĐỂ TRỐNG 
    "Chưa thanh toán"  // R: Trạng thái
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
  
  // SePay có thể gửi nội dung ở trường content hoặc gateway_content
  const noiDungCK = (data.content || data.description || "").toUpperCase();
  const soTienCK = data.transferAmount || data.amount || 0;
  
  for (let i = values.length - 1; i > 0; i--) {
    let phoneTrongSheet = values[i][3].toString().replace(/\D/g, ''); // Cột D (SĐT)
    
    // Bỏ qua nếu dòng đó không có số điện thoại
    if (phoneTrongSheet === "") continue;
    
    // Kiểm tra nếu nội dung Chuyển khoản có chứa số điện thoại
    if (noiDungCK.indexOf(phoneTrongSheet) !== -1) {
      // Ghi trạng thái đã thanh toán vào cột R (Cột thứ 18)
      sheet.getRange(i + 1, 18).setValue("✅ ĐÃ THANH TOÁN: " + soTienCK.toLocaleString() + "đ");
      
      // Flush để đảm bảo dữ liệu được ghi ngay lập tức vào sheet
      SpreadsheetApp.flush();
      
      return ContentService.createTextOutput(JSON.stringify({success: true})).setMimeType(ContentService.MimeType.JSON);
    }
  }
  return ContentService.createTextOutput(JSON.stringify({success: false, message: "Không tìm thấy số điện thoại trong nội dung CK"}));
}
