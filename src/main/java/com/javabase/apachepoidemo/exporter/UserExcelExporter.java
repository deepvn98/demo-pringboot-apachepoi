package com.javabase.apachepoidemo.exporter;

import com.javabase.apachepoidemo.model.User;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.List;

public class UserExcelExporter {
    private XSSFWorkbook workbook;
    private XSSFSheet sheet;
    private List<User> listUsers;

    public UserExcelExporter() {
    }

    public UserExcelExporter(List<User> listUsers) {
        this.listUsers = listUsers;
        workbook = new XSSFWorkbook();
        sheet = workbook.createSheet("User");

    }

    private void createCell(Row row, int columsCount, Object value, CellStyle style){
        sheet.autoSizeColumn(columsCount);
        Cell cell = row.createCell(columsCount);
        if (value instanceof Long){
            cell.setCellValue((Long)value);
        } if (value instanceof Integer){
            cell.setCellValue((Integer)value);
        } if (value instanceof String){
            cell.setCellValue((String)value);
        } if (value instanceof Boolean){
            cell.setCellValue((Boolean) value);
        }
        cell.setCellStyle(style);
    }
    private void writeHeaderRow() {
        List<String> propertyUser = User.getProperty();
        Row row = sheet.createRow(0);
        CellStyle style = workbook.createCellStyle();
        XSSFFont font = workbook.createFont();
        font.setFontHeight(18);
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setFont(font);
        createCell(row,0,"User Information",style);
        sheet.addMergedRegion(new CellRangeAddress(0,0,0,3));
        font.setFontHeight(14);
        style.setFont(font);
        Row row1 = sheet.createRow(1);
        for (int i = 0; i < propertyUser.size(); i++) {
            createCell(row1,i,propertyUser.get(i),style);
        }
    }
    private void writeDataRow(List<User> listUsers) {
        int rowCount =2;
        CellStyle style = workbook.createCellStyle();
        XSSFFont font = workbook.createFont();
        font.setFontHeight(14);
        style.setFont(font);
        for (User user: listUsers){
            int columnCount = 0;
            Row row = sheet.createRow(rowCount++);
            createCell(row,columnCount++,user.getId(),style);
            createCell(row,columnCount++,user.getName(),style);
            createCell(row,columnCount++,user.getUsername(),style);
            createCell(row,columnCount++,user.getEmail(),style);
        }

    }

    public void export(HttpServletResponse response, List<User> listUsers) {
        DateFormat dateFormat = new SimpleDateFormat("yyy-MM-dd_HH:mm:ss");
        String currentDateTime = dateFormat.format(new Date());
        String fileName = "user "+currentDateTime+".xlsx";
        response.setContentType("application/octet-stream");
        String headerKey = "Content-Disposition";
        String headerValue = "attachment; filename="+fileName;
        response.setHeader(headerKey,headerValue);
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









    //    Cach 1
//    public UserExcelExporter(List<User> listUsers) {
//        this.workbook = new XSSFWorkbook();
//        this.sheet = workbook.createSheet("User");
//        this.listUsers = listUsers;
//    }
//
//    private void writeHeaderRow() {
//        List<String> propertyUser = User.getProperty();
//        Row row = sheet.createRow(0);
//        Cell cell;
//        for (int i = 0; i < propertyUser.size(); i++) {
//            cell = row.createCell(i);
//            cell.setCellValue(propertyUser.get(i));
//        }
//    }
//
//    private void writeDataRow(List<User> listUsers) {
//        List<String> propertyUser = User.getProperty();
//        Row row;
//        Cell cell;
//        for (int i = 0; i < listUsers.size(); i++) {
//            row = sheet.createRow(i + 1);
//            cell = row.createCell(0);
//            cell.setCellValue(listUsers.get(i).getId());
//            cell = row.createCell(1);
//            cell.setCellValue(listUsers.get(i).getName());
//            cell = row.createCell(2);
//            cell.setCellValue(listUsers.get(i).getUsername());
//            cell = row.createCell(3);
//            cell.setCellValue(listUsers.get(i).getEmail());
//        }
//
//    }
//
//    public void export(HttpServletResponse response, List<User> listUsers) {
//        DateFormat dateFormat = new SimpleDateFormat("yyy-MM-dd_HH:mm:ss");
//        String currentDateTime = dateFormat.format(new Date());
//        String fileName = "user "+currentDateTime+".xlsx";
//        response.setContentType("application/octet-stream");
//        String headerKey = "Content-Disposition";
//        String headerValue = "attachment; filename="+fileName;
//        response.setHeader(headerKey,headerValue);
//        writeHeaderRow();
//        writeDataRow(listUsers);
//        ServletOutputStream outputStream = null;
//        try {
//            outputStream = response.getOutputStream();
//            workbook.write(outputStream);
//            workbook.close();
//            outputStream.close();
//        } catch (IOException e) {
//            e.printStackTrace();
//        }
//    }
}
