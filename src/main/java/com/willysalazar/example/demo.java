package com.willysalazar.example;

import org.apache.poi.ss.usermodel.Cell;
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
import java.util.concurrent.TimeUnit;


public class demo {

    private static int rowNumber = 1; // Starting row number in Excel

    public static void main(String[] args) throws IOException {
        // List of names and corresponding companies
        /*String[][] profiles = {
                {"Brendan Achariyakosol,Evolute Capital", "Evolute Capital"},
        };*/
        try {
            FileInputStream file = new FileInputStream(new File("C:\\Users\\rbolla\\Documents\\TestData.xlsx"));
            XSSFWorkbook workbook = new XSSFWorkbook(file);
            XSSFSheet sheet = workbook.getSheet("Sheet1");

            // Check if the sheet exists
            if (sheet == null) {
                System.out.println("Sheet with name 'Sheet1' not found in the workbook.");
                // Handle this situation as needed (e.g., close resources and exit)
                file.close();
                workbook.close();
                return;
            }

            // Get the number of rows in the sheet
            int rowCount = sheet.getLastRowNum() - sheet.getFirstRowNum();

            // Create an array to store profiles
            String[][] profiles = new String[rowCount][2]; // Assuming you want to store Full Name and Company

            // Loop through each row and get "Full Name" and "Company" columns
            for (int i = 1; i <= rowCount; i++) {
                Row row = sheet.getRow(i);

                Cell fullNameCell = row.getCell(5); // Assuming "Full Name" is in the 6th column (0-based index)
                Cell companyCell = row.getCell(6); // Assuming "Company" is in the 7th column (0-based index)

                String fullName = (fullNameCell != null) ? fullNameCell.getStringCellValue() : "";
                String company = (companyCell != null) ? companyCell.getStringCellValue() : "";

                profiles[i - 1][0] = fullName;
                profiles[i - 1][1] = company;
            }

            file.close();
            workbook.close();

            System.setProperty("webdriver.chrome.driver", "D:\\ChromeDriver\\119_chromedriver-win64 (1)\\chromedriver-win64\\chromedriver.exe");
            WebDriver driver = new ChromeDriver();
            WebDriverWait wait = new WebDriverWait(driver, 10);

            Workbook outputWorkbook = new XSSFWorkbook();
            Sheet outputSheet = outputWorkbook.createSheet("LinkedInProfiles");

            Row headerRow = outputSheet.createRow(0);
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

            for (String[] profile : profiles) {
                String fullName = profile[0];
                String company = profile[1];

                // Split the full name into first name and last name
                String[] nameParts = fullName.split("\\s+");
                String firstName = nameParts[0];
                String lastName = nameParts[nameParts.length - 1];

                // Navigate to LinkedIn home page to ensure we start with a clean search
                driver.get("https://www.linkedin.com");

                // Wait for the search box to be present
                WebElement searchBox = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//input[@placeholder='Search']")));

                // Clear the search input before entering a new name
                searchBox.clear();

                // Construct a more specific search query using first name, last name, and company
                String searchQuery = String.format("%s %s %s", firstName, lastName, company);
                searchBox.sendKeys(searchQuery);
                searchBox.sendKeys(Keys.RETURN);

                // Check if the "View full profile" button is present
                try {
                    WebElement viewFullProfileButton = wait.until(
                            ExpectedConditions.elementToBeClickable(By.xpath("//span[text()='View full profile']")));

                    viewFullProfileButton.click();

                    // Introduce a delay to give LinkedIn time to load the profile page
                    TimeUnit.SECONDS.sleep(5);

                    // Scroll down to load all the content on the profile page
                    JavascriptExecutor js = (JavascriptExecutor) driver;
                    js.executeScript("window.scrollTo(0, document.body.scrollHeight);");

                    String currentUrl = driver.getCurrentUrl();
                    String username = extractUsernameFromUrl(currentUrl);
                    System.out.println("Username extracted from URL: " + username);

                    // Write to Excel sheet
                    writeToExcel(outputSheet, rowNumber++, fullName, username);

                } catch (Exception e) {
                    // Write to Excel sheet for unsuccessful run
                    writeToExcel(outputSheet, rowNumber++, fullName, "Failed");
                }

                // Go back to the search results page
                driver.navigate().back();
            }

            // Save the workbook to a file
            try (FileOutputStream fileOut = new FileOutputStream("LinkedInProfiles.xlsx")) {
                outputWorkbook.write(fileOut);
            } catch (Exception e) {
                e.printStackTrace();
            }

            driver.quit();

        } catch (IOException e) {
            e.printStackTrace();
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
