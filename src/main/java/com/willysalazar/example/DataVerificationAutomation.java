package com.willysalazar.example;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.*;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;
import java.util.NoSuchElementException;
import java.util.concurrent.TimeUnit;

public class DataVerificationAutomation {
    public static void main(String[] args) {
        long startTime = System.currentTimeMillis();
        System.setProperty("webdriver.chrome.driver", "D:\\ChromeDriver\\119_chromedriver-win64 (1)\\chromedriver-win64\\chromedriver.exe");
        WebDriver driver = new ChromeDriver();
        WebDriverWait wait = new WebDriverWait(driver, 10);

        // Navigate to LinkedIn login page
        driver.get("https://www.linkedin.com/login");

        // Enter username and password
        WebElement usernameInput = driver.findElement(By.id("username"));
        WebElement passwordInput = driver.findElement(By.id("password"));

        usernameInput.sendKeys("rohithbolla97@gmail.com");
        passwordInput.sendKeys("Encoress@123");

        // Click on the login button
        WebElement loginButton = driver.findElement(By.xpath("//button[@type='submit']"));
        loginButton.click();

        try (FileInputStream fileInputStream = new FileInputStream("C:\\Users\\rbolla\\Documents\\TestData.xlsx");
             Workbook workbook = new XSSFWorkbook(fileInputStream);
             FileOutputStream fileOut = new FileOutputStream("LinkedInProfiles.xlsx")) {

            Sheet sheet = workbook.getSheet("Sheet1");

            // Create header row
            Row headerRow = sheet.getRow(0);
            String[] headers = {"Sno", "Name", "Company_name", "Status"};
            for (int i = 0; i < headers.length; i++) {
                Cell cell = headerRow.createCell(i);
                cell.setCellValue(headers[i]);
            }

            int rowNum = 1;
            int endRow = 2;

            for (int i = 1; i <= endRow; i++) {
                Row row = sheet.getRow(i);
                String name = row.getCell(5).getStringCellValue();
                String company = row.getCell(6).getStringCellValue();

                System.out.println("name"+name);
                System.out.println("Company"+company);

                // Step 2: Search for the name and press Enter
                WebElement searchBox = driver.findElement(By.xpath("//input[@placeholder='Search']"));

                // Clear the search input before entering a new name
                searchBox.clear();

                searchBox.sendKeys(name);
                searchBox.sendKeys(Keys.RETURN);

                // Check if the "View full profile" button is present
                try {
                    WebElement viewFullProfileButton = wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//span[text()='View full profile']")));
                    viewFullProfileButton.click();
                } catch (TimeoutException e) {
                    // If the "View full profile" button is not present, click on the link corresponding to the name
                    try {
                        WebElement nameLink = wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//span[text()='" + name + "']/ancestor::a")));
                        nameLink.click();
                    } catch (TimeoutException ex) {
                        System.out.println("No search result found for " + name);
                        continue; // Move to the next iteration if the link is not found
                    }
                }

                // Check the current company name on the profile
                try {
                    WebElement companyNameElement = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//span[contains(text(), '" + company + "')]/ancestor::span")));
                    WebElement nextLineElement = companyNameElement.findElement(By.xpath("./following-sibling::*"));

                    // Check if the text contains "Present" or "present"
                    String status;
                    if (nextLineElement.getText().toLowerCase().contains("present")) {
                        status = "Same company";
                    } else {
                        status = "Different company";
                    }

                    // Write data to Excel sheet
                    Row resultRow = sheet.createRow(rowNum++);
                    resultRow.createCell(0).setCellValue(rowNum - 1);
                    resultRow.createCell(1).setCellValue(name);
                    resultRow.createCell(2).setCellValue(company);
                    resultRow.createCell(3).setCellValue(status);
                } catch (TimeoutException e) {
                    System.out.println(name + " works in a different company");
                    // Write data to Excel sheet for profiles with different organizations
                    Row resultRow = sheet.createRow(rowNum++);
                    resultRow.createCell(0).setCellValue(rowNum - 1);
                    resultRow.createCell(1).setCellValue(name);
                    resultRow.createCell(2).setCellValue(company);
                    resultRow.createCell(3).setCellValue("Different company");
                }

                // Go back to the search results page
                driver.navigate().back();
            }

            // Save the workbook to a file
            workbook.write(fileOut);
        } catch (IOException e) {
            e.printStackTrace();
        }

        // Close the browser
        driver.quit();
        long endTime = System.currentTimeMillis();
        long executionTime = endTime - startTime;
        System.out.println("Execution time: " + executionTime + " milliseconds");

    }
}
