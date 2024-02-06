package com.willysalazar.example;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.*;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;

import java.io.FileOutputStream;
import java.net.URI;
import java.net.URISyntaxException;
import java.util.concurrent.TimeUnit;

public class LinkedinScraper {

    private static int rowNumber = 1; // Starting row number in Excel

    public static void main(String[] args) {
        // List of names and corresponding companies
        String[][] profiles = {
                {"Kenneth Bloom,Preferred Compounding Corp.", "Preferred Compounding Corp."},
                {"Stephen Blumenreich,BNP Paribas", "BNP Paribas"},
        };

        System.setProperty("webdriver.chrome.driver", "D:\\ChromeDriver\\121 version\\chromedriver-win64\\chromedriver.exe");

        WebDriver driver = new ChromeDriver();
        WebDriverWait wait = new WebDriverWait(driver, 10);

        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("LinkedInProfiles");

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

        for (String[] profile : profiles) {
            String name = profile[0];

            // Step 2: Search for the name and press Enter
            WebElement searchBox = driver.findElement(By.xpath("//input[@placeholder='Search']"));

            // Clear the search input before entering a new name
            searchBox.clear();

            searchBox.sendKeys(name);
            searchBox.sendKeys(Keys.RETURN);
            String[] parts = name.split(",");
            name = parts[0];
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
                    System.out.println("Username extracted from URL: " + username);

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
                            .elementToBeClickable(By.xpath("//span[text()='" + name + "']/ancestor::a")));
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
                    System.out.println("No search result found for " + name);
                    continue;
                }
            }
            // Go back to the search results page
            driver.navigate().back();
        }

        // Save the workbook to a file
        try (FileOutputStream fileOut = new FileOutputStream("LinkedInProfiles.xlsx")) {
            workbook.write(fileOut);
        } catch (Exception e) {
            e.printStackTrace();
        }

        driver.quit();
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
