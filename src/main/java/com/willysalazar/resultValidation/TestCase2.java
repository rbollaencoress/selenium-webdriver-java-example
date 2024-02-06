package com.willysalazar.resultValidation;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

public class TestCase2 {
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
        int updatedJobTitleColumnIndex1 = -1; // New column index for Updated Job Title in Sheet1
        int updatedCompanyColumnIndex1 = -1; // New column index for Updated Company in Sheet1
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
            } else if ("Updated Job Title".equalsIgnoreCase(cellValue)) {
                updatedJobTitleColumnIndex1 = i;
            } else if ("Updated Company".equalsIgnoreCase(cellValue)) {
                updatedCompanyColumnIndex1 = i;
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

        if (updatedCompanyColumnIndex1 == -1) {
            throw new IllegalArgumentException("Updated Company column not found in Sheet1.");
        }

        if (updatedJobTitleColumnIndex1 == -1) {
            throw new IllegalArgumentException("Updated Job Title column not found in Sheet1.");
        }

        if (presentCompanyColumnIndex2 == -1) {
            throw new IllegalArgumentException("Present Company column not found in Sheet2.");
        }

        if (designationColumnIndex2 == -1) {
            throw new IllegalArgumentException("Designation column not found in Sheet2.");
        }

        for (int i = 1; i < sheet1RowCount; i++) {
            Row row1 = excelSheet1.getRow(i);
            Cell companyCell = row1.getCell(companyColumnIndex1);
            Cell jobTitleCell = row1.getCell(jobTitleColumnIndex1);
            String companyName = companyCell != null ? companyCell.getStringCellValue() : null;
            String jobTitle = jobTitleCell != null ? jobTitleCell.getStringCellValue() : null;
            boolean companyMatchFound = false;
            boolean jobTitleMatchFound = false;
            String presentCompany = "";
            String designation = "";

            if (companyName != null && jobTitle != null) {
                for (int j = 1; j < sheet2RowCount; j++) {
                    Row row2 = excelSheet2.getRow(j);
                    Cell presentCompanyCell = row2.getCell(presentCompanyColumnIndex2);
                    Cell designationCell = row2.getCell(designationColumnIndex2);
                    presentCompany = presentCompanyCell != null ? presentCompanyCell.getStringCellValue() : null;
                    designation = designationCell != null ? designationCell.getStringCellValue() : null;

                    if (presentCompany != null && designation != null) {
                        if (companyName.equalsIgnoreCase(presentCompany)) {
                            companyMatchFound = true;
                        }

                        if (jobTitle.equalsIgnoreCase(designation)) {
                            jobTitleMatchFound = true;
                        }

                        if (companyMatchFound && jobTitleMatchFound) {
                            break; // Break the loop if both matches are found
                        }
                    }
                }
            }

            Cell checkCell = row1.createCell(checkColumnIndex1);
            if (companyMatchFound && jobTitleMatchFound) {
                checkCell.setCellValue("Ok");
            } else if (companyMatchFound) {
                checkCell.setCellValue("Changed Job Title");
                if (updatedJobTitleColumnIndex1 >= 0) {
                    Cell updatedJobTitleCell = row1.createCell(updatedJobTitleColumnIndex1);
                    updatedJobTitleCell.setCellValue(designation); // Set the designation as updated job title
                }
            } else if (jobTitleMatchFound) {
                checkCell.setCellValue("Changed Company");
                if (updatedCompanyColumnIndex1 >= 0) {
                    Cell updatedCompanyCell = row1.createCell(updatedCompanyColumnIndex1);
                    updatedCompanyCell.setCellValue(presentCompany); // Set the present company as updated company
                }
            } else {
                checkCell.setCellValue("Changed Company and Job Title");
                if (updatedJobTitleColumnIndex1 >= 0) {
                    Cell updatedJobTitleCell = row1.createCell(updatedJobTitleColumnIndex1);
                    updatedJobTitleCell.setCellValue(designation); // Set the designation as updated job title
                }
                if (updatedCompanyColumnIndex1 >= 0) {
                    Cell updatedCompanyCell = row1.createCell(updatedCompanyColumnIndex1);
                    updatedCompanyCell.setCellValue(presentCompany); // Set the present company as updated company
                }
            }
        }

        FileOutputStream outputStream = new FileOutputStream("D:\\Contacts\\ResultsValidation.xlsx");
        excelWorkbook.write(outputStream);
        excelWorkbook.close();
        outputStream.close();
    }
}
