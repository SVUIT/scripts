function onFormSubmit(e) {
  // Lấy các mục phản hồi từ biểu mẫu
  var responses = e.values;
  
  // Xác định vị trí của câu trả lời chứa liên kết tải lên tệp
  // Thông thường, e.values[0] là thời gian gửi, e.values[1], e.values[2], ... là các câu trả lời
  var fileColumnIndex = 3; // Thay đổi giá trị này theo đúng vị trí của cột chứa URL tải lên
  
  // Lấy URL của tệp đã tải lên
  var fileUrls = responses[fileColumnIndex];

  // Trích xuất các URL (nếu nhiều tệp, mỗi URL sẽ được phân tách bằng dấu phẩy)
  var urls = fileUrls.split(',');

  // Duyệt qua tất cả các URL tệp đã tải lên
  for (var i = 0; i < urls.length; i++) {
    // Lấy ID của tệp từ URL
    var fileId = urls[i].match(/[-\w]{25,}/);
    
    if (fileId) {
      var file = DriveApp.getFileById(fileId[0]);
      var fileName = file.getName();
      
      // Xử lý tên file để xóa phần tên người dùng sau dấu "-"
      var newName = fileName.replace(/ - .+\./, '.'); // Xóa phần tên sau dấu "-" trước phần mở rộng file
      
      // Đổi tên file
      file.setName(newName);
    }
  }
}
