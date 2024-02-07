package com.willysalazar.resultValidation;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import java.io.FileInputStream;
import java.io.FileOutputStream;

public class TestCase2_3rdTime {
    public static void main(String[] args) {
        try {
            // Load the Excel file
            FileInputStream fis = new FileInputStream("D:\\Contacts\\ResultsValidation.xlsx");
            Workbook workbook = WorkbookFactory.create(fis);
            Sheet sheet1 = workbook.getSheet("sheet1");
            Sheet sheet2 = workbook.getSheet("sheet2");

            // Iterate through rows and compare Company and Job Title
            for (int i = 1; i <= sheet1.getLastRowNum(); i++) {
                Row row1 = sheet1.getRow(i);
                Row row2 = sheet2.getRow(i);

                String company1 = row1.getCell(0).getStringCellValue();
                String presentCompany = row2.getCell(0).getStringCellValue();
                String jobTitle1 = row1.getCell(1).getStringCellValue();
                String designation = row2.getCell(1).getStringCellValue();

                // Compare Company and set Check column accordingly
                if (company1.equals(presentCompany) && jobTitle1.equals(designation)) {
                    row1.createCell(2).setCellValue("OK");
                } else {
                    row1.createCell(2).setCellValue("Change Company");
                    // Update Updated Company and Job Title columns
                    row1.createCell(3).setCellValue(presentCompany);
                    row1.createCell(4).setCellValue(designation);
                }
            }

            // Write the changes back to the Excel file
            FileOutputStream fos = new FileOutputStream("D:\\Contacts\\ResultsValidation.xlsx");
            workbook.write(fos);

            // Close resources
            fos.close();
            workbook.close();
            fis.close();

            System.out.println("Comparison completed successfully.");
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
