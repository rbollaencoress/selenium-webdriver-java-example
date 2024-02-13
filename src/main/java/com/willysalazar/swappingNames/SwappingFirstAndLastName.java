package com.willysalazar.swappingNames;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.*;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.net.URI;
import java.net.URISyntaxException;
import java.util.Arrays;
import java.util.Scanner;
import java.util.concurrent.TimeUnit;

public class SwappingFirstAndLastName {
    public static void main(String[] args) throws IOException{

            long startTime = System.currentTimeMillis();

            Scanner sc = new Scanner(System.in);
            System.out.println("Enter the row count in input Excel");
            int maxRows = sc.nextInt();
            String username = null;
            try {
                //FileInputStream inputStream = new FileInputStream(new File("C:\\Users\\rbolla\\Desktop\\MyTestData.xlsx"));
                FileInputStream inputStream = new FileInputStream(new File("D:\\Contacts\\MyData.xlsx"));
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
               // int fullNameColumnIndex = -1;
                int companyColumnIndex = 7;

                int firstNameColumnIndex = -1;
                int middleNameColumnIndex = -1;
                int lastNameColumnIndex = -1;

                Row headerRow1 = excelSheet.getRow(headerRowNum);
                if (headerRow1 != null) {
                    // Iterating through the cells in the header row to find the "Full Name" column
                    for (int i = headerRow1.getFirstCellNum(); i < headerRow1.getLastCellNum(); i++) {
                        String cellValue = headerRow1.getCell(i).getStringCellValue();
                        if ("First Name".equalsIgnoreCase(cellValue)) {
                            firstNameColumnIndex = i;
                        } else if ("Middle Name".equalsIgnoreCase(cellValue)) {
                            middleNameColumnIndex = i;
                        } else if ("Last Name".equalsIgnoreCase(cellValue)) {
                            lastNameColumnIndex = i;
                        }
                    }
                }

                // Check if the "Full Name" and "Company" columns are found
                if (firstNameColumnIndex == -1 || middleNameColumnIndex == -1 || lastNameColumnIndex == -1) {
                    System.out.println("Columns 'First Name', 'Middle Name' and 'Last name' not found in the sheet.");
                    inputStream.close();
                    excelWorkbook.close();
                    return;
                }
                System.setProperty("webdriver.chrome.driver", "D:\\ChromeDriver\\121 version\\chromedriver-win64\\chromedriver.exe");

                WebDriver driver = new ChromeDriver();
                WebDriverWait wait = new WebDriverWait(driver, 3);

                ChromeOptions options = new ChromeOptions();
                options.addArguments("--incognito");

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

                usernameInput.sendKeys("ahladhguptagg27@gmail.com");
                passwordInput.sendKeys("Qwerty2@");

                WebElement loginButton = driver.findElement(By.xpath("//button[@type='submit']"));
                loginButton.click();

                int rowNumber = 1; // Start from row 1 (after the header row)

                // Iterate through rows and get profiles from Excel sheet
                for (int rowNum = headerRowNum + 1; rowNum <= excelSheet.getLastRowNum() && rowNumber <= maxRows; rowNum++) {
                    Row row = excelSheet.getRow(rowNum);

                    if (row != null) {
                        String lastName = "";
                        String middleName = "";
                        String firstName = "";
                        String company = "";

                        // Fetching values from the cells
                        Cell lastNameCell = row.getCell(lastNameColumnIndex);
                        if (lastNameCell != null) {
                            lastName = lastNameCell.getStringCellValue();
                        }

                        Cell middleNameCell = row.getCell(middleNameColumnIndex);
                        if (middleNameCell != null) {
                            middleName = middleNameCell.getStringCellValue();
                        }

                        Cell firstNameCell = row.getCell(firstNameColumnIndex);
                        if (firstNameCell != null) {
                            firstName = firstNameCell.getStringCellValue();
                        }

                        if (companyColumnIndex >= 0) { // Ensure companyColumnIndex is valid
                            Cell companyCell = row.getCell(companyColumnIndex);
                            if (companyCell != null) {
                                company = companyCell.getStringCellValue();
                            }
                        }

                        String[] fullName = {lastName," ",middleName," ",firstName};
                        String[] profile = {lastName+" "+middleName+" "+firstName +","+company};
                        System.out.println("Company:"+company);
                        System.out.println("Arrays:"+Arrays.asList(profile));
                        WebElement searchBox = driver.findElement(By.xpath("//input[@placeholder='Search']"));
                        searchBox.clear();

                        searchBox.sendKeys(profile);
                        searchBox.sendKeys(Keys.RETURN);
                        try {

                            WebElement viewFullProfileButton = wait.until(
                                    ExpectedConditions.elementToBeClickable(By.xpath("//span[text()='View full profile']")));
                            try {
                                viewFullProfileButton.click();

                                // Introduce a delay to give LinkedIn time to load the profile page
                                TimeUnit.SECONDS.sleep(3);

                                // Scroll down to load all the content on the profile page
                                JavascriptExecutor js = (JavascriptExecutor) driver;
                                js.executeScript("window.scrollTo(0, document.body.scrollHeight);");

                                String currentUrl = driver.getCurrentUrl();
                                username = extractUsernameFromUrl(currentUrl);
                                System.out.println("fullName:"+Arrays.asList(fullName)+" username:"+ username);

                                // Write to Excel sheet
                                writeToExcel(sheet, rowNumber++, profile[0], username);

                            } catch (Exception e) {
                                // Write to Excel sheet for unsuccessful run
                                writeToExcel(sheet, rowNumber++, profile[0], "No Results Found");
                                System.out.println("fullName:"+fullName+" username:No Results Found");
                                continue;
                            }
                        } catch (TimeoutException e) {
//                                writeToExcel(sheet, rowNumber++, profile[0], "No Records Found");
//                                System.out.println("fullName:"+ fullName+"hyperlink username:No Records Found");
//                                continue;
                                searchWithDifferentCases(profile, searchBox);
                                writeToExcel(sheet, rowNumber++, profile[0], username);
                                continue;

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

    private static void searchWithDifferentCases(String[] profile, WebElement searchBox) {
        System.out.println(Arrays.asList(profile));
        String[] splitName = profile[0].split(" ");
        String[] searchQueries = new String[4];
        if (splitName.length >= 3) {
            // Case 1: middle name first name company
            searchQueries[0] = splitName[1] + " " + splitName[2] + " " + profile[0];

            // Case 2: last name first name company
            searchQueries[1] = splitName[0] + " " + splitName[2] + " " + profile[0];

            // Case 3: last name first name company
            searchQueries[2] = splitName[0] + " " + splitName[1] + " " + profile[0];

            // Case 4: first name middle name company (if middle name is not blank)
            if (!splitName[1].isEmpty()) {
                searchQueries[3] = splitName[2] + " " + splitName[1] + " " + profile[0];
            }
        }

        // Search with different queries
        for (String query : searchQueries) {
            if (query != null) {
                searchBox.clear();
                searchBox.sendKeys(query);
                searchBox.sendKeys(Keys.RETURN);
                // Add a wait time between searches
                try {
                    Thread.sleep(2000); // Wait for 2 seconds between searches
                } catch (InterruptedException e) {
                    e.printStackTrace();
                }
                // Add further logic here to handle the search results
                // You can reuse the existing logic or modify it as needed
            }
        }
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
                for (int i = 0; i < pathSegments.length - 1; i++) {
                    if ("in".equals(pathSegments[i])) {
                        return pathSegments[i + 1];
                    }
                }
            } catch (URISyntaxException e) {
                e.printStackTrace();
            }

            return null;
        }
}
