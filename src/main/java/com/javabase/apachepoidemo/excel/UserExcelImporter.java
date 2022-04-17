package com.javabase.apachepoidemo.excel;

import com.javabase.apachepoidemo.model.User;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

public class UserExcelImporter {
    public List<User> excelImport(String pathFile) {
        List<User> userList = new ArrayList<>();

        Long id = null;
        String name = "";
        String username = "";
        String email = "";
        FileInputStream fileInputStream;
        try {
            fileInputStream = new FileInputStream(pathFile);
            Workbook workbook = new XSSFWorkbook(fileInputStream);
            Sheet sheet = workbook.getSheetAt(0);
            int rows = sheet.getLastRowNum(); //  get all row number
            int cols = sheet.getRow(1).getLastCellNum();// get all column number
            for (int r = 1; r <= rows; r++) {
                Row row = sheet.getRow(r);
                for (int c = 0; c < cols; c++) {
                    Cell cell = row.getCell(c);
                    switch (cell.getColumnIndex()) {
                        case 0:
                            id = (long) cell.getNumericCellValue();
                            System.out.println(id);

                            break;
                        case 1:
                            name = cell.getStringCellValue();
                            System.out.println(name);

                            break;
                        case 2:
                            username = cell.getStringCellValue();
                            System.out.println(username);

                            break;
                        case 3:
                            email = cell.getStringCellValue();
                            System.out.println(email);
                            break;

                    }

                }
                User user = new User(id, name, username, email);
                System.out.println();
                userList.add(user);
            }
            workbook.close();
        } catch (Exception e) {
            e.printStackTrace();
        }

        return userList;

    }


}
