package io.nspai;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.formula.functions.Column;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbookType;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

public class App {

    private static String[] columns = {"FirstName", "LastName", "Email", "Date Of Birth"};



    public static void main(String[] args) throws IOException {

        List<Contact> contacts = new ArrayList<>();
        contacts.add(new Contact("Gauthier","Ninespace","gauthier@test.com","12/12/1900"));
        contacts.add(new Contact("Josh","Ninespace","josh@test.com","12/12/1910"));
        contacts.add(new Contact("Jemima","Ninespace","jemima@test.com","12/12/1920"));
        contacts.add(new Contact("John","Ninespace","john@test.com","12/12/1930"));
        contacts.add(new Contact("Grace","Ninespace","grace@test.com","12/12/1940"));

        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Contacts");
        Font headerFont = workbook.createFont();
        headerFont.setBold(true);
        headerFont.setFontHeightInPoints((short) 17);
        headerFont.setColor(IndexedColors.RED.getIndex());

        CellStyle headerCellStyle = workbook.createCellStyle();
        headerCellStyle.setFont(headerFont);

        Row headerRow = sheet.createRow(0);

        //create header cells and add them to the header row

        for (int i = 0; i < columns.length; i++) {
            Cell cell = headerRow.createCell(i);
            cell.setCellValue(columns[i]);
            cell.setCellStyle(headerCellStyle);
        }

        int rowNum = 1;

        for (Contact contact: contacts) {
            Row row = sheet.createRow(rowNum++);
            row.createCell(0).setCellValue(contact.getFirstName());
            row.createCell(1).setCellValue(contact.getLastName());
            row.createCell(2).setCellValue(contact.getEmail());
            row.createCell(3).setCellValue(contact.getDateOfBirth());
        }

        for (int i = 0; i < columns.length; i++) {
            sheet.autoSizeColumn(1);
        }

        try (
                FileOutputStream fileOutputStream = new FileOutputStream("Contacts.xlsx")) {
            workbook.write(fileOutputStream);
        }
        workbook.close();
    }
}
