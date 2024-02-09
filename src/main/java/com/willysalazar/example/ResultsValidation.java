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

        Row headerRow1 = excelSheet1.getRow(0);
        Row headerRow2 = excelSheet2.getRow(0);

        int companyColumnIndex1 = -1;
        int jobTitleColumnIndex1 = -1;
        int checkColumnIndex1 = -1;
        int presentCompanyColumnIndex2 = -1;
        int designationColumnIndex2 = -1;

        for (int i = headerRow1.getFirstCellNum(); i < headerRow1.getLastCellNum(); i++) {
            String cellValue = headerRow1.getCell(i).getStringCellValue();
            if ("Company".equalsIgnoreCase(cellValue)) {
                companyColumnIndex1 = i;
            } else if ("Job Title".equalsIgnoreCase(cellValue)) {
                jobTitleColumnIndex1 = i;
            } else if ("Check".equalsIgnoreCase(cellValue)) {
                checkColumnIndex1 = i;
            }
        }

        for (int i = headerRow2.getFirstCellNum(); i < headerRow2.getLastCellNum(); i++) {
            String cellValue = headerRow2.getCell(i).getStringCellValue();
            if ("Present Company".equalsIgnoreCase(cellValue)) {
                presentCompanyColumnIndex2 = i;
            } else if ("Designation".equalsIgnoreCase(cellValue)) {
                designationColumnIndex2 = i;
            }
        }

        if (presentCompanyColumnIndex2 == -1) {
            throw new IllegalArgumentException("Present Company column not found in Sheet2.");
        }

        if (designationColumnIndex2 == -1) {
            throw new IllegalArgumentException("Designation column not found in Sheet2.");
        }

//        int presentCompanyColumnIndex1 = headerRow1.getLastCellNum();
//        headerRow1.createCell(presentCompanyColumnIndex1).setCellValue("Present Company from Sheet2");
//        int designationColumnIndex1 = headerRow1.getLastCellNum();
//        headerRow1.createCell(designationColumnIndex1).setCellValue("Designation from Sheet2");
        int presentCompanyColumnIndex1 = findOrCreateColumn(headerRow1, "Present Company from Sheet2");
        int designationColumnIndex1 = findOrCreateColumn(headerRow1, "Designation from Sheet2");
        System.out.println(sheet1RowCount+"sheet1RowCount");
        for (int i = 1; i < sheet1RowCount; i++) {
            Row row1 = excelSheet1.getRow(i);
            //String companyName = row1.getCell(companyColumnIndex1).getStringCellValue();
            String companyName;
            Cell cell = row1.getCell(companyColumnIndex1);
            if (cell != null) {
                companyName = cell.getStringCellValue();
            } else {
                companyName = ""; // Assign empty string if the cell is null
            }
            System.out.println(companyName);
            //String jobTitle = row1.getCell(jobTitleColumnIndex1).getStringCellValue();
            String jobTitle;
            Cell jobTitleCell = row1.getCell(jobTitleColumnIndex1);
            if (jobTitleCell != null) {
                jobTitle = jobTitleCell.getStringCellValue();
            } else {
                jobTitle = ""; // Assign empty string if the cell is null
            }
            boolean companyMatchFound = false;
            boolean jobTitleMatchFound = false;

            // Set Check cell to "Not Found" if Company is null
            if (companyName == null || companyName.trim().isEmpty()) {
                Cell checkCell = row1.createCell(checkColumnIndex1);
                Row row2 = excelSheet2.getRow(i); // Adjusted row index to match Sheet1
                String presentCompany = row2.getCell(presentCompanyColumnIndex2).getStringCellValue();
                String designation = row2.getCell(designationColumnIndex2).getStringCellValue();

                row1.createCell(presentCompanyColumnIndex1).setCellValue(presentCompany);
                row1.createCell(designationColumnIndex1).setCellValue(designation);
                checkCell.setCellValue("Added new company and Job title");
                continue; // Skip to the next row
            }

            for (int j = 1; j < sheet2RowCount; j++) {
                Row row2 = excelSheet2.getRow(j);
                String presentCompany = row2.getCell(presentCompanyColumnIndex2).getStringCellValue();
                String designation = row2.getCell(designationColumnIndex2).getStringCellValue();

                if (companyName.equalsIgnoreCase(presentCompany)) {
                    companyMatchFound = true;
                }

                if (jobTitle.equalsIgnoreCase(designation)) {
                    jobTitleMatchFound = true;
                }

                if (companyMatchFound && jobTitleMatchFound) {
                    break;
                }
            }

            if (!companyMatchFound || !jobTitleMatchFound) {
                Row row2 = excelSheet2.getRow(i); // Adjusted row index to match Sheet1
                String presentCompany = row2.getCell(presentCompanyColumnIndex2).getStringCellValue();
                String designation = row2.getCell(designationColumnIndex2).getStringCellValue();

                row1.createCell(presentCompanyColumnIndex1).setCellValue(presentCompany);
                row1.createCell(designationColumnIndex1).setCellValue(designation);
            }

            Cell checkCell = row1.createCell(checkColumnIndex1);
            if (companyMatchFound && jobTitleMatchFound) {
                checkCell.setCellValue("Ok");
            } else if (companyMatchFound) {
                checkCell.setCellValue("Changed Job Title");
            } else if (jobTitleMatchFound) {
                checkCell.setCellValue("Changed Company");
            } else {
                checkCell.setCellValue("Changed Company and Job Title");
            }
        }

        FileOutputStream outputStream = new FileOutputStream("D:\\Contacts\\ResultsValidation.xlsx");
        excelWorkbook.write(outputStream);
        excelWorkbook.close();
        outputStream.close();
    }
    private static int findOrCreateColumn(Row headerRow, String columnName) {
        for (int i = headerRow.getFirstCellNum(); i < headerRow.getLastCellNum(); i++) {
            String cellValue = headerRow.getCell(i).getStringCellValue();
            if (columnName.equalsIgnoreCase(cellValue)) {
                return i;
            }
        }
        // If the column doesn't exist, create it
        int columnIndex = headerRow.getLastCellNum();
        headerRow.createCell(columnIndex).setCellValue(columnName);
        return columnIndex;
    }
}
