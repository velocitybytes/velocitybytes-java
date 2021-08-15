package com.velocitybytes.excelutils;

import org.apache.poi.ss.usermodel.*;

import java.io.File;
import java.io.IOException;
import java.util.Iterator;

/**
 * Use:
 * ExcelReader excelReader = new ExcelReader();
 * try {
 *     excelReader.excelReader("data/velocitybytes.xlsx");
 * } catch (IOException e) {
 *     e.printStackTrace();
 * }
 */
public class ExcelReader {

    public void excelReader(String file) throws IOException {

        // Create a Workbook from the given Excel file
        Workbook workbook = WorkbookFactory.create(new File(file));

        // Get number of sheets in the Workbook
        System.out.println("Workbook has " + workbook.getNumberOfSheets() + " sheet(s).");

        /*
         * Iterate over all the sheets in the Workbook
         */
        // 1. Obtain sheetIterator and iterate over it
        Iterator<Sheet> sheetIterator = workbook.sheetIterator();
        System.out.println("Sheets:");
        while (sheetIterator.hasNext()) {
            Sheet sheet = sheetIterator.next();
            System.out.println("-> " + sheet.getSheetName());
        }

        // 2. Using for-each loop
        for (Sheet sheet: workbook) {
            System.out.println("-> " + sheet.getSheetName());
        }

        // 3. Using Java 8 forEach loop with lambda
        workbook.forEach(sheet -> System.out.println(sheet.getSheetName()));

        /*
         * Iterate over rows and columns in Sheet
         */

        // Get the sheet at required index
        Sheet sheet = workbook.getSheetAt(0);

        // Create a DataFormatter to format and get each cell's value as String
        DataFormatter dataFormatter = new DataFormatter();

        // 1. Obtain rowIterator and columnIterator
        Iterator<Row> rowIterator = sheet.rowIterator();
        while (rowIterator.hasNext()) {
            Row row = rowIterator.next();

            // Iterate over columns of current row
            Iterator<Cell> cellIterator = row.cellIterator();
            while (cellIterator.hasNext()) {
                Cell cell = cellIterator.next();
                String cellValue = dataFormatter.formatCellValue(cell);
                System.out.println(cellValue);
            }
        }

        // 2. Use for-each loop
        for (Row row: sheet) {
            for (Cell cell: row) {
                String cellValue = dataFormatter.formatCellValue(cell);
                System.out.println(cellValue);
            }
        }

        // 3. Java 8 forEach loop with lambda
        sheet.forEach(row -> row.forEach(cell -> {
            String cellValue = dataFormatter.formatCellValue(cell);
            System.out.println(cellValue);
        }));

        // Close the Workbook
        workbook.close();
    }
}
