package com.javabase.apachepoidemo.Controller;

import com.javabase.apachepoidemo.excel.UserExcelExporter;
import com.javabase.apachepoidemo.excel.UserExcelImporter;
import com.javabase.apachepoidemo.model.User;
import com.javabase.apachepoidemo.repository.UserRepository;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.HttpStatus;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import javax.servlet.http.HttpServletResponse;
import java.util.List;

@RestController
@RequestMapping("/api/user")
public class UserController {
    @Autowired
    private UserRepository userRepository;

    @GetMapping("/getall/user")
    public ResponseEntity<?> getAllUser() {

        return new ResponseEntity<>(userRepository.findAllUser(), HttpStatus.OK);
    }

    @GetMapping("/user/export")
    public void exportToExcel(HttpServletResponse response) {
        List<User> listUsers = userRepository.findAllUser();
        UserExcelExporter userExcelExporter = new UserExcelExporter(listUsers);
        userExcelExporter.export(response, listUsers);
    }

    @GetMapping("/user/import")
    public String importToExcel() {
        String filepath = "C:\\Users\\admin\\Downloads\\user 2022-04-17_00_38_42.xlsx";
        UserExcelImporter userExcelImporter = new UserExcelImporter();
//        userExcelImporter.excelImport(filepath);
        List<User> listUsers = userExcelImporter.excelImport(filepath);
        userRepository.saveAll(listUsers);
        return "import Successfully!";
    }

}
