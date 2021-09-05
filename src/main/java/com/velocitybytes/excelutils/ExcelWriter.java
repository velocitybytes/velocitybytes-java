package com.velocitybytes.excelutils;

import com.velocitybytes.model.Person;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

public class ExcelWriter {

    private static final String[] columns = {"Name", "Username", "Email"};

    private static final List<Person> persons = new ArrayList<>();

    static {
        persons.add(new Person("Srivastava", "srivastava", "srivastava@gmail.com"));
        persons.add(new Person("VelocityBytes", "velocitybytes", "srivastava@gmail.com"));
        persons.add(new Person("Killer Frost", "killerfrost", "killerfrost@gmail.com"));
    }

    public void excelWriter(String file) throws IOException {
        // Create a workbook - for .xlsx files
        Workbook workbook = new XSSFWorkbook(); // HSSFWorkbook() for .xls files

        /* CreationHelper helps us create instances of various things like DataFormat,
           Hyperlink, RichTextString etc., in a format (HSSF, XSSF) independent way
        */
        CreationHelper creationHelper = workbook.getCreationHelper();

        // Create a Sheet
        Sheet sheet = workbook.createSheet("Person");

        // Create a Font for styling header cells
        Font headerFont = workbook.createFont();
        headerFont.setBold(false);
        headerFont.setFontHeightInPoints((short) 12);
        headerFont.setColor(IndexedColors.GREEN.getIndex());

        // Create CellStyle with font
        CellStyle cellStyle = workbook.createCellStyle();
        cellStyle.setFont(headerFont);

        // Create a Row
        Row headerRow = sheet.createRow(0);

        // Create cells
        for (int i = 0; i < columns.length; i++) {
            Cell cell = headerRow.createCell(i);
            cell.setCellValue(columns[i]);
            cell.setCellStyle(cellStyle);
        }

        int rowNum = 1;

        for (Person person: persons) {
            Row row = sheet.createRow(rowNum++);
            row.createCell(0).setCellValue(person.getName());
            row.createCell(1).setCellValue(person.getUsername());
            row.createCell(2).setCellValue(person.getEmail());
        }

        // Resize all columns to fit the content size
        for (int i = 0; i < columns.length; i++) {
            sheet.autoSizeColumn(i);
        }

        // Write the output to a file
        FileOutputStream outputStream = new FileOutputStream(file);
        workbook.write(outputStream);
        outputStream.close();

        // Close the workbook
        workbook.close();
    }
}
