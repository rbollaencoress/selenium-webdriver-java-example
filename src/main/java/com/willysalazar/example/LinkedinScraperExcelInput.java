package com.willysalazar.example;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.*;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.net.URI;
import java.net.URISyntaxException;
import java.util.Scanner;
import java.util.concurrent.TimeUnit;


public class LinkedinScraperExcelInput {
    public static void main(String[] args) {

        long startTime = System.currentTimeMillis();

        Scanner sc = new Scanner(System.in);
        System.out.println("Enter the row count in input Excel");
        int maxRows = sc.nextInt();
        try {
            FileInputStream inputStream = new FileInputStream(new File("C:\\Users\\rbolla\\Desktop\\MyTestData.xlsx"));
            XSSFWorkbook excelWorkbook = new XSSFWorkbook(inputStream);
            XSSFSheet excelSheet = excelWorkbook.getSheet("Sheet1");

            // Check if the sheet exists
            if (excelSheet == null) {
                System.out.println("Sheet with name 'Sheet1' not found in the workbook.");
                inputStream.close();
                excelWorkbook.close();
                return;
            }

            int headerRowNum = 0; // Assuming the header is in the first row

            // Finding the column index for "Full Name"
            int fullNameColumnIndex = -1;
            int companyColumnIndex = -1;
            org.apache.poi.ss.usermodel.Row headerRow1 = excelSheet.getRow(headerRowNum);
            if (headerRow1 != null) {
                // Iterating through the cells in the header row to find the "Full Name" column
                for (int i = headerRow1.getFirstCellNum(); i < headerRow1.getLastCellNum(); i++) {
                    String cellValue = headerRow1.getCell(i).getStringCellValue();
                    if ("Full Name".equalsIgnoreCase(cellValue)) {
                        fullNameColumnIndex = i;
                    } else if ("Company".equalsIgnoreCase(cellValue)) {
                        companyColumnIndex = i;
                    }
                }
            }

            // Check if the "Full Name" and "Company" columns are found
            if (fullNameColumnIndex == -1 || companyColumnIndex == -1) {
                System.out.println("Columns 'Full Name' and 'Company' not found in the sheet.");
                inputStream.close();
                excelWorkbook.close();
                return;
            }

            WebDriver driver = new ChromeDriver();
            WebDriverWait wait = new WebDriverWait(driver, 10);

            Workbook workbook = new XSSFWorkbook();
            Sheet sheet = workbook.createSheet("LinkedInProfiles_usernames");

            Row headerRow = sheet.createRow(0);
            headerRow.createCell(0).setCellValue("Sno");
            headerRow.createCell(1).setCellValue("Name");
            headerRow.createCell(2).setCellValue("Username");

            // Navigate to LinkedIn login page
            driver.get("https://www.linkedin.com/login");

            // Enter username and password
            WebElement usernameInput = driver.findElement(By.id("username"));
            WebElement passwordInput = driver.findElement(By.id("password"));

            usernameInput.sendKeys("rohithbolla97@gmail.com");
            passwordInput.sendKeys("Encoress@123");

            WebElement loginButton = driver.findElement(By.xpath("//button[@type='submit']"));
            loginButton.click();

            int rowNumber = 1; // Start from row 1 (after the header row)

            // Iterate through rows and get profiles from Excel sheet
            for (int rowNum = headerRowNum + 1; rowNum <= excelSheet.getLastRowNum() && rowNumber <= maxRows; rowNum++) {
                Row row = excelSheet.getRow(rowNum);

                if (row != null) {
                    String fullName = row.getCell(fullNameColumnIndex).getStringCellValue();
                    String company = row.getCell(companyColumnIndex).getStringCellValue();
                    String[] profile = {fullName, ","+company};
                    // Step 2: Search for the name and press Enter
                    WebElement searchBox = driver.findElement(By.xpath("//input[@placeholder='Search']"));

                    // Clear the search input before entering a new name
                    searchBox.clear();

                    searchBox.sendKeys(profile);
                    searchBox.sendKeys(Keys.RETURN);
                    String[] parts = fullName.split(",");
                    fullName = parts[0];
                    // Check if the "View full profile" button is present
                    try {
                        WebElement viewFullProfileButton = wait.until(
                                ExpectedConditions.elementToBeClickable(By.xpath("//span[text()='View full profile']")));
                        try {
                            viewFullProfileButton.click();

                            // Introduce a delay to give LinkedIn time to load the profile page
                            TimeUnit.SECONDS.sleep(5);

                            // Scroll down to load all the content on the profile page
                            JavascriptExecutor js = (JavascriptExecutor) driver;
                            js.executeScript("window.scrollTo(0, document.body.scrollHeight);");

                            String currentUrl = driver.getCurrentUrl();
                            String username = extractUsernameFromUrl(currentUrl);
                            System.out.println("Username extracted for" +fullName+" from URL: " + username);

                            // Write to Excel sheet
                            writeToExcel(sheet, rowNumber++, profile[0], username);

                        } catch (Exception e) {
                            // Write to Excel sheet for unsuccessful run
                            writeToExcel(sheet, rowNumber++, profile[0], "No Results `Found`");
                            continue;
                        }
                    } catch (TimeoutException e) {
                        // If the "View full profile" button is not present, click on the link
                        // corresponding to the name
                        try {
                            WebElement nameLink = wait.until(ExpectedConditions
                                    .elementToBeClickable(By.xpath("//span[text()='" + fullName + "']/ancestor::a")));
                            nameLink.click();

                            // Introduce a delay to give LinkedIn time to load the profile page
                            TimeUnit.SECONDS.sleep(5);

                            // Scroll down to load all the content on the profile page
                            JavascriptExecutor js = (JavascriptExecutor) driver;
                            js.executeScript("window.scrollTo(0, document.body.scrollHeight);");

                            String currentUrl = driver.getCurrentUrl();
                            String username = extractUsernameFromUrl(currentUrl);
                            System.out.println("Username extracted from URL: " + username);

                            // Write to Excel sheet
                            writeToExcel(sheet, rowNumber++, profile[0], username);

                        } catch (TimeoutException | InterruptedException ex) {
                            // Write to Excel sheet for unsuccessful run
                            writeToExcel(sheet, rowNumber++, profile[0], "No Results foundl");
                            System.out.println("No search result found for " + fullName);
                            continue;
                        }
                    }
                    // Go back to the search results page
                    driver.navigate().back();
                }
            }

            // Save the workbook to a file
            try (FileOutputStream fileOut = new FileOutputStream("LinkedInProfiles.xlsx")) {
                workbook.write(fileOut);
            } catch (Exception e) {
                e.printStackTrace();
            }

            driver.quit();

        } catch (IOException e) {
            e.printStackTrace();
        }
        long endTime = System.currentTimeMillis();
        long elapsedTimeInMillis = endTime - startTime;
        long elapsedTimeInSeconds = elapsedTimeInMillis / 1000;
        long elapsedTimeInMinutes = elapsedTimeInSeconds / 60;

        System.out.println("Total execution time: " + elapsedTimeInMinutes + " minutes and " +
                (elapsedTimeInSeconds % 60) + " seconds");
    }

    private static void writeToExcel(Sheet sheet, int rowNum, String name, String username) {
        Row row = sheet.createRow(rowNum);
        row.createCell(0).setCellValue(rowNum);
        row.createCell(1).setCellValue(name);
        row.createCell(2).setCellValue(username);
    }

    private static String extractUsernameFromUrl(String url) {
        try {
            URI uri = new URI(url);
            String path = uri.getPath();

            // Extract the username from the path
            String[] pathSegments = path.split("/");
            for (int i = 0; i < pathSegments.length; i++) {
                if ("in".equals(pathSegments[i]) && i + 1 < pathSegments.length) {
                    return pathSegments[i + 1];
                }
            }
        } catch (URISyntaxException e) {
            e.printStackTrace();
        }

        return null;
    }
}
