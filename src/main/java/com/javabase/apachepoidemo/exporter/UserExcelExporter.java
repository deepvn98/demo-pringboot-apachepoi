package com.javabase.apachepoidemo.exporter;

import com.javabase.apachepoidemo.model.User;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.util.List;

public class UserExcelExporter {
    private XSSFWorkbook workbook;
    private XSSFSheet sheet;
    private List<User> listUsers;

    public UserExcelExporter() {
    }

    public UserExcelExporter(List<User> listUsers) {
        this.workbook = new XSSFWorkbook();
        this.sheet = workbook.createSheet("User");
        this.listUsers = listUsers;
    }

    private void writeHeaderRow() {
        List<String> propertyUser = User.getProperty();
        Row row = sheet.createRow(0);
        Cell cell;
        for (int i = 0; i < propertyUser.size(); i++) {
            cell = row.createCell(i);
            cell.setCellValue(propertyUser.get(i));
        }
    }

    private void writeDataRow(List<User> listUsers) {
        List<String> propertyUser = User.getProperty();
        Row row;
        Cell cell;
        for (int i = 0; i < listUsers.size(); i++) {
            row = sheet.createRow(i + 1);
            cell = row.createCell(0);
            cell.setCellValue(listUsers.get(i).getId());
            cell = row.createCell(1);
            cell.setCellValue(listUsers.get(i).getName());
            cell = row.createCell(2);
            cell.setCellValue(listUsers.get(i).getUsername());
            cell = row.createCell(3);
            cell.setCellValue(listUsers.get(i).getEmail());
        }

    }

    public void export(HttpServletResponse response, List<User> listUsers) {
        writeHeaderRow();
        writeDataRow(listUsers);
        ServletOutputStream outputStream = null;
        try {
            outputStream = response.getOutputStream();
            workbook.write(outputStream);
            workbook.close();
            outputStream.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
