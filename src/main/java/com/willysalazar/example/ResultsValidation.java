package com.willysalazar.example;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

public class ResultsValidation {
    public static void main(String[] args) throws IOException {
        FileInputStream inputStream = new FileInputStream(new File("D:\\Contacts\\ResultsValidation.xlsx"));
        XSSFWorkbook excelWorkbook = new XSSFWorkbook(inputStream);
        XSSFSheet excelSheet1 = excelWorkbook.getSheet("Sheet1");
        XSSFSheet excelSheet2 = excelWorkbook.getSheet("Sheet2");

        int sheet1RowCount = excelSheet1.getPhysicalNumberOfRows();
        int sheet2RowCount = excelSheet2.getPhysicalNumberOfRows();

        // Assuming header row is the first row
        Row headerRow1 = excelSheet1.getRow(0);
        Row headerRow2 = excelSheet2.getRow(0);

        int companyColumnIndex1 = -1;
        int checkColumnIndex1 = -1;
        int presentCompanyColumnIndex2 = -1;
        int previousCompaniesColumnIndex2 = -1;

        // Find column indices in Sheet1 for "Company" and "Check"
        for (int i = headerRow1.getFirstCellNum(); i < headerRow1.getLastCellNum(); i++) {
            String cellValue = headerRow1.getCell(i).getStringCellValue();
            if ("Company".equalsIgnoreCase(cellValue)) {
                companyColumnIndex1 = i;
            } else if ("Check".equalsIgnoreCase(cellValue)) {
                checkColumnIndex1 = i;
            }
        }

        // Check if "Company" column was found
        if (companyColumnIndex1 == -1) {
            throw new IllegalArgumentException("Company column not found in Sheet1.");
        }

        // Check if "Check" column was found
        if (checkColumnIndex1 == -1) {
            throw new IllegalArgumentException("Check column not found in Sheet1.");
        }

        // Find column index in Sheet2 for "Present Company"
        for (int i = headerRow2.getFirstCellNum(); i < headerRow2.getLastCellNum(); i++) {
            String cellValue = headerRow2.getCell(i).getStringCellValue();
            if ("Present Company".equalsIgnoreCase(cellValue)) {
                presentCompanyColumnIndex2 = i;
                break; // Exit the loop once "Present Company" column is found
            }
        }

        // Check if "Present Company" column was found
        if (presentCompanyColumnIndex2 == -1) {
            throw new IllegalArgumentException("Present Company column not found in Sheet2.");
        }

        // Find column index in Sheet2 for "Previous Companies"
        for (int i = headerRow2.getFirstCellNum(); i < headerRow2.getLastCellNum(); i++) {
            String cellValue = headerRow2.getCell(i).getStringCellValue();
            if ("Previous Companies".equalsIgnoreCase(cellValue)) {
                previousCompaniesColumnIndex2 = i;
                break; // Exit the loop once "Previous Companies" column is found
            }
        }

        // Check if "Previous Companies" column was found
        if (previousCompaniesColumnIndex2 == -1) {
            throw new IllegalArgumentException("Previous Companies column not found in Sheet2.");
        }

        // Iterate through rows in Sheet1
        for (int i = 1; i < sheet1RowCount; i++) {
            Row row1 = excelSheet1.getRow(i);
            String companyName = row1.getCell(companyColumnIndex1).getStringCellValue();
            boolean matchFound = false;

            // Iterate through rows in Sheet2 to check present company
            for (int j = 1; j < sheet2RowCount; j++) {
                Row row2 = excelSheet2.getRow(j);
                String presentCompany = row2.getCell(presentCompanyColumnIndex2).getStringCellValue();

                if (companyName.equalsIgnoreCase(presentCompany)) {
                    matchFound = true;
                    break;
                }
            }

            Cell checkCell = row1.createCell(checkColumnIndex1);
            checkCell.setCellValue(matchFound ? "Ok" : "Changed Company");
        }

        // Save the changes back to the Excel file
        FileOutputStream outputStream = new FileOutputStream("D:\\Contacts\\ResultsValidation.xlsx");
        excelWorkbook.write(outputStream);
        excelWorkbook.close();
        outputStream.close();
    }
}
