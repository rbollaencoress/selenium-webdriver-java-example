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
            //int presentCompanyColumnIndex1 = getColumnIndex(excelSheet1, "Present company");
            int presentCompanyColumnIndex1 = 5;
            int designationColumnIndex1 = getColumnIndex(excelSheet1, "Designation");


            int sNoColumnIndex2 = getColumnIndex(excelSheet2, "SNo");
            int fullNameColumnIndex2 = getColumnIndex(excelSheet2, "Full Name");
            int companyColumnIndex2 = getColumnIndex(excelSheet2, "Company");
            int jobTitleColumnIndex2 = getColumnIndex(excelSheet2, "Job Title");
            int checkColumnIndex = getColumnIndex(excelSheet2, "Check"); // Assuming "Check" column exists in SRC sheet

            for (int i = 1; i <= rowCount; i++) { // Assuming the first row is header
                Row row1 = excelSheet1.getRow(i);
                Row row2 = excelSheet2.getRow(i);
                Row headerRow1 = excelSheet1.getRow(0);
                Row headerRow2 = excelSheet2.getRow(0);

                // Get cell values based on column indexes
                String sNo1 = getCellValueAsString(row1.getCell(sNoColumnIndex1));
                String fullName1 = row1.getCell(fullNameColumnIndex1).getStringCellValue();
                //String presentCompany1 = row1.getCell(presentCompanyColumnIndex1).getStringCellValue();
                String presentCompany1 = "";
                Cell presentCompanyCell1 = row1.getCell(presentCompanyColumnIndex1);
                if (presentCompanyCell1 != null) {

                    presentCompany1 = presentCompanyCell1.getStringCellValue();
                }
                //String designation1 = row1.getCell(designationColumnIndex1).getStringCellValue();
                String designation1;
                Cell designationCell1 = row1.getCell(designationColumnIndex1);
                if (designationCell1 != null) {
                    designation1 = designationCell1.getStringCellValue();
                } else {
                    designation1 = ""; // Assign empty string if the cell is null
                }


                String sNo2 = getCellValueAsString(row2.getCell(sNoColumnIndex2));
                String fullName2 = row2.getCell(fullNameColumnIndex2).getStringCellValue();
                //String company2 = row2.getCell(companyColumnIndex2).getStringCellValue();
                String company2;
                Cell companyCell2 = row2.getCell(companyColumnIndex2);
                if (companyCell2 != null) {
                    company2 = companyCell2.getStringCellValue();
                } else {
                    company2 = ""; // Assign empty string if the cell is null
                }
                //String jobTitle2 = row2.getCell(jobTitleColumnIndex2).getStringCellValue();
                String jobTitle2 = "";
                Cell jobTitleCell2 = row2.getCell(jobTitleColumnIndex2);
                if (jobTitleCell2 != null) {
                    jobTitle2 = jobTitleCell2.getStringCellValue();
                }

                if (sNo1.equals(sNo2) && fullName1.equals(fullName2)) {
                    if((presentCompany1.equals("") || designation1.equals("")) && (company2.equals("") || jobTitle2.equals(""))){
                        row2.createCell(checkColumnIndex).setCellValue("Not Found");
                    }
                    else if (presentCompany1.equals(company2) && designation1.equals(jobTitle2)) {
                        // Both present company and designation match, set "OK" in the "Check" column
                        row2.createCell(checkColumnIndex).setCellValue("OK");
                    }
                    else if(company2.equals("") || jobTitle2.equals("")) {
                        if(!(presentCompany1.equals("") || designation1.equals(""))){
                            row2.createCell(checkColumnIndex).setCellValue("Added new company and Job Title");
                            row2.createCell(companyColumnIndex2).setCellValue(presentCompany1);
                            row2.createCell(jobTitleColumnIndex2).setCellValue(designation1);
                        }
                    }
                    else if((presentCompany1.equals("") || designation1.equals("")) && !(company2.equals("") || (jobTitle2.equals("")))){
                        row2.createCell(checkColumnIndex).setCellValue("Not Found");
                    }
                    else if(!presentCompany1.equals(company2)){
                        row2.createCell(checkColumnIndex).setCellValue("Changed Company");
                        row2.createCell(companyColumnIndex2).setCellValue(presentCompany1);
                        row2.createCell(jobTitleColumnIndex2).setCellValue(designation1);
                    } else if (!designation1.equals(jobTitle2)) {
                        row2.createCell(checkColumnIndex).setCellValue("Changed Job Title");
                        row2.createCell(jobTitleColumnIndex2).setCellValue(designation1);
                    }
//                        else {
//                        // Either present company or designation doesn't match, update company in SRC sheet
//                        row2.createCell(checkColumnIndex).setCellValue("Added new company and Job Title");
//                        row2.createCell(companyColumnIndex2).setCellValue(presentCompany1);
//                        row2.createCell(jobTitleColumnIndex2).setCellValue(designation1);
//                    }
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
