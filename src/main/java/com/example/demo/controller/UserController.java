package com.example.demo.controller;


import com.example.demo.Repository.UserRepository;
import com.example.demo.etity.User;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.HttpHeaders;
import org.springframework.http.HttpStatus;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;


@RestController
//@RequestMapping("/api/users")
public class UserController {


    @Autowired
    private UserRepository userRepository;

//    api 1 :  Tải về 1 file excel mẫu chứa các trường thông tin như hình 1 bên dưới

    @GetMapping("/download-sample-excel")
    public ResponseEntity<byte[]> downloadSampleExcel() throws IOException {
        //XSSFWorkbook: Đây là lớp từ thư viện Apache POI được dùng để tạo và quản lý một workbook (file Excel) mới với định dạng .xlsx.
        //Sheet : Tạo một sheet mới trong workbook với tên là "Mẫu Thông Tin". Mỗi file Excel có thể có nhiều sheet, và sheet này chứa dữ liệu người dùng
        // Tạo một workbook và sheet mới cho file Excel
        XSSFWorkbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Mẫu Thông Tin");

        // Row: Tạo hàng đầu tiên (header) chứa các trường thông tin
        Row headerRow = sheet.createRow(0);
        String[] columns = {"ID", "Tên", "Tuổi", "Địa chỉ", "Email"};

        for (int i = 0; i < columns.length; i++) {
            Cell cell = headerRow.createCell(i);
            cell.setCellValue(columns[i]);
        }

//        ByteArrayOutputStream để ghi dữ liệu của file Excel vào một mảng byte.
//        workbook.write(outputStream);: Ghi toàn bộ nội dung của workbook (file Excel) vào đối tượng outputStream.
//        workbook.close();: Đóng workbook sau khi hoàn thành thao tác ghi dữ liệu để giải phóng tài nguyên.
        // Tạo file Excel dưới dạng byte array để gửi về client
        ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
        workbook.write(outputStream);
        workbook.close();

        byte[] bytes = outputStream.toByteArray();

        // Thiết lập header cho phản hồi
        HttpHeaders headers = new HttpHeaders();
        headers.add("Content-Disposition", "attachment; filename=sample.xlsx");

        return new ResponseEntity<>(bytes, headers, HttpStatus.OK);
    }


//    api 2 :  API import excel   -  nhập thông tin user vào file mẫu ở trên  thực hiện import - Thông tin user lưu vào bảng user

    @PostMapping("/import-users")
    public ResponseEntity<String> importUsersFromExcel(@RequestParam("file") MultipartFile file) {
        if (file.isEmpty()) {
            return new ResponseEntity<>("File không được để trống", HttpStatus.BAD_REQUEST);
        }

        List<User> users = new ArrayList<>();

        try (InputStream inputStream = file.getInputStream()) {
            Workbook workbook = new XSSFWorkbook(inputStream);
            Sheet sheet = workbook.getSheetAt(0); // Lấy sheet đầu tiên

            for (int i = 1; i <= sheet.getLastRowNum(); i++) { // Bắt đầu từ hàng thứ 2 (bỏ qua tiêu đề)
                Row row = sheet.getRow(i);

                // Kiểm tra xem hàng có tồn tại không
                if (row == null) {
                    continue; // Bỏ qua hàng nếu nó không tồn tại
                }

                User user = new User();
                user.setAge((int) row.getCell(0).getNumericCellValue()); // id
                user.setName(row.getCell(1).getStringCellValue()); // Tên
                user.setAge((int) row.getCell(2).getNumericCellValue()); // Tuổi
                user.setAddress(row.getCell(3).getStringCellValue()); // Địa chỉ
                user.setEmail(row.getCell(4).getStringCellValue()); // Email

                users.add(user); // Thêm người dùng vào danh sách
            }

            // Lưu thông tin người dùng vào cơ sở dữ liệu
            userRepository.saveAll(users);
            return new ResponseEntity<>("Import thành công", HttpStatus.OK);

        } catch (IOException e) {
            return new ResponseEntity<>("Lỗi khi đọc file: " + e.getMessage(), HttpStatus.INTERNAL_SERVER_ERROR);
        } catch (Exception e) {
            return new ResponseEntity<>("Đã xảy ra lỗi: " + e.getMessage(), HttpStatus.BAD_REQUEST);
        }
    }

//    api 3 : API export excel - Xuất thông tin trong database thành file excel

    @GetMapping("/export-users")
    public ResponseEntity<byte[]> exportUsersToExcel() throws IOException {
        // Tạo workbook và sheet
        XSSFWorkbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Danh sách người dùng");

        // Tạo hàng đầu tiên (header)
        Row headerRow = sheet.createRow(0);
        String[] columns = {"ID", "Tên", "Tuổi", "Địa chỉ", "Email"};
        for (int i = 0; i < columns.length; i++) {
            Cell cell = headerRow.createCell(i);
            cell.setCellValue(columns[i]);
        }

        // Lấy danh sách người dùng từ cơ sở dữ liệu
        List<User> users = userRepository.findAll();

        // Điền dữ liệu người dùng vào các hàng tiếp theo
        int rowNum = 1;
        for (User user : users) {
            Row row = sheet.createRow(rowNum++);

            row.createCell(0).setCellValue(user.getId());
            row.createCell(1).setCellValue(user.getName());
            row.createCell(2).setCellValue(user.getAge());
            row.createCell(3).setCellValue(user.getAddress());
            row.createCell(4).setCellValue(user.getEmail());
        }

        // Ghi workbook vào byte array
        ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
        workbook.write(outputStream);
        workbook.close();

        byte[] excelBytes = outputStream.toByteArray();

        // Thiết lập header để trả về file Excel
        HttpHeaders headers = new HttpHeaders();
        headers.add("Content-Disposition", "attachment; filename=users.xlsx");

        return new ResponseEntity<>(excelBytes, headers, HttpStatus.OK);
    }
}





