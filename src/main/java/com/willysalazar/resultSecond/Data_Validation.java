package com.willysalazar.resultSecond;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;


public class Data_Validation {
    public static void main(String[] args) {
        try {
            FileInputStream inputStream = new FileInputStream(new File("D:\\Contacts\\DataValidation\\ResultDataValidaion.xlsx"));
            XSSFWorkbook excelWorkbook = new XSSFWorkbook(inputStream);
            XSSFSheet excelSheet1 = excelWorkbook.getSheet("EXT");
            XSSFSheet excelSheet2 = excelWorkbook.getSheet("SRC");

            int rowCount = Math.min(excelSheet1.getLastRowNum(), excelSheet2.getLastRowNum());

            // Get column indexes based on column names
            int sNoColumnIndex1 = getColumnIndex(excelSheet1, "SNo");
            int fullNameColumnIndex1 = getColumnIndex(excelSheet1, "Full Name");
            int sNoColumnIndex2 = getColumnIndex(excelSheet2, "SNo");
            int fullNameColumnIndex2 = getColumnIndex(excelSheet2, "Full Name");
            int checkColumnIndex = getColumnIndex(excelSheet2, "Check"); // Assuming "Check" column exists in SRC sheet

            for (int i = 1; i <= rowCount; i++) { // Assuming the first row is header
                Row row1 = excelSheet1.getRow(i);
                Row row2 = excelSheet2.getRow(i);

                // Get cell values based on column indexes
                String sNo1 = getCellValueAsString(row1.getCell(sNoColumnIndex1));
                String fullName1 = row1.getCell(fullNameColumnIndex1).getStringCellValue();

                String sNo2 = getCellValueAsString(row2.getCell(sNoColumnIndex2));
                String fullName2 = row2.getCell(fullNameColumnIndex2).getStringCellValue();

                if (sNo1.equals(sNo2) && fullName1.equals(fullName2)) {
                    
                }
                else if (!sNo1.equals(sNo2) || !fullName1.equals(fullName2)) {
                    // Set the value of the "Check" column to "Not matching Full name"
                    row2.createCell(checkColumnIndex).setCellValue("Not matching S No or Full name");
                }
            }

            // Write the changes back to the workbook
            FileOutputStream outputStream = new FileOutputStream(new File("D:\\Contacts\\DataValidation\\ResultDataValidaion.xlsx"));
            excelWorkbook.write(outputStream);

            excelWorkbook.close();
            inputStream.close();
            outputStream.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private static int getColumnIndex(XSSFSheet sheet, String columnName) {
        Row headerRow = sheet.getRow(0);
        if (headerRow != null) {
            for (Cell cell : headerRow) {
                if (cell.getStringCellValue().equals(columnName)) {
                    return cell.getColumnIndex();
                }
            }
        }
        return -1; // Column not found
    }

    private static String getCellValueAsString(Cell cell) {
        if (cell == null) {
            return "";
        }

        if (cell.getCellType() == CellType.NUMERIC) {
            return String.valueOf((int) cell.getNumericCellValue());
        } else {
            return cell.getStringCellValue();
        }
    }
}
