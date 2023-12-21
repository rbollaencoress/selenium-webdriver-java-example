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
    public static void main(String[] args)  {
        // List of names and corresponding companies
        long startTime = System.currentTimeMillis();

        String[][] profiles = {
                {"Brendan Achariyakosol", "Evolute Capital"},
                {"Bob Adams", "Growth Operators"},
                {"Brad Adams", "TM Capital Corporation"},
                {"Bryan Adams", "Integrity Marketing Group"},
                {"Jane Adams", "Piper Jaffray & Company"},
                {"Katherine Adams", "JETNET"},
                {"Mark Affolter", "Ares Management LLC"},
                {"Sergey Agafonkin", "Opus Financial Partners"},
                {"Donald Agee", "CMF Associates"},
                {"Rahul Aggarwal", "Brentwood Associates"},
                {"Steven Aguiar", "Provident Healthcare Partners"},
                {"Mike Ahern", "Medsphere Systems Corporation"},
                {"Randy Ahlm", "Boston Retail Solutions"},
                {"Giray Aikar", "Convey Health Solutions"},
                {"Brad Akason", "Lincoln International"},
                {"Omar Akbar", "Aperion Management"},
                {"Jim Akerhielm", "iVision"},
                {"Darren Alcus", "Capital One"},
                {"Mohammed Algazi", "Wells Fargo"},
                {"Jeremy Alinder", "Blue Ops"},
                {"James Allegretti", "Deloitte & Touche LLP"},
                {"Scott Allen", "Morris Manning & Martin, LLP"},
                {"Howard Allred", "West Star Aviation, Inc."},
                {"Mike Altman", "Cortland Partners"},
                {"Lee Alvarez", "Marwood Group"},
                {"Scott Ames", "Cascadia Capital"},
                {"Richard Amistadi", "Amistadi Associates LLC"},
                {"Road Ammons", "ECD Capital Partners"},
                {"Elliot Amundson", "TripleTree"},
                {"Andy Anderson", "Thirdera"},
                {"Bill Anderson", "HealthTech"},
                {"Brent Anderson", "Ernst & Young"},
                {"Chad Anderson", "JetCraft Corporation"},
                {"Keith Anderson", "Piper Jaffray & Company"},
                {"Lain Anderson", "L.E.K. Consulting"},
                {"Robert Andrews", "Robert W. Baird"},
                {"Thomas Anspach", "Anspach"},
                {"Gregg Antenen", "Vistio"},
                {"Anita Antenucci", "3Wire Partners"},
                {"Michael Anton", "LogixHealth"},
                {"Kojo Appenteng", "Stifel Financial Corp."},
                {"Tony Aquilina", "Thomas H. Lee Partners, L.P."},
                {"Robert Arditi", "Bessemer Venture Partners"},
                {"Steve Arentsen", "Wells Fargo"},
                {"Alan Aria", "Capital One"},
                {"Paul Arne", "Morris Manning & Martin, LLP"},
                {"Robert Arsov", "Hoplon Capital LLC"},
                {"Jeff Ash", "Summit Hosting - Secure Cloud Hosting"},
                {"Cullen Atchison", "The Houston Group Realty Advisors"},
                {"Kevin Atchue", "TM Capital Corporation"},
                {"Stephen Avery", "Ventra Health"},
                {"Amit Aysola", "Create Health Ventures / Wanxiang America Corporation"},
                {"Gene Babcock", "MarketLab"},
                {"Kim Babcock", "Surgical Information Systems"},
                {"Christopher Babick", "Monroe Capital LLC"},
                {"Samuel Bachman", "Northland Capital Markets"},
                {"Azad Badakhsh", "Moelis & Company"},
                {"Blair Badham", "Lafayette Square"},
                {"Benedict Baerst", "Aquiline Capital Partners LLC"},
                {"David Bahk", "Jefferies & Company, Inc."},
                {"Seth Bair", "Keefe, Bruyette & Woods"},
                {"Bill Baker", "KPMG"},
                {"Colin Baker", "Stellus Capital Management, LLC"},
                {"Carter Balfour", "Norwest Mezzanine Partners"},
                {"Marvin Banks", "Cortland Partners"},
                {"Brice Baradel", "Highview Capital"},
                {"Deborah Barbosa", "SitelogIQ"}
        };

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

        // Create a new Excel workbook
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("LinkedIn Profiles");
        try {
            // Create a new sheet
            // Create header row
            Row headerRow = sheet.createRow(0);
            String[] headers = { "Sno", "Name", "Company_name", "Status" };
            for (int i = 0; i < headers.length; i++) {
                Cell cell = headerRow.createCell(i);
                cell.setCellValue(headers[i]);
            }

            int rowNum = 1;

            for (String[] profile : profiles) {
                String name = profile[0];
                String company = profile[1];

                // Step 2: Search for the name and press Enter
                WebElement searchBox = driver.findElement(By.xpath("//input[@placeholder='Search']"));

                // Clear the search input before entering a new name
                searchBox.clear();

                searchBox.sendKeys(name);
                searchBox.sendKeys(Keys.RETURN);

                // Check if the "View full profile" button is present
                try {
                    WebElement viewFullProfileButton = wait.until(
                            ExpectedConditions.elementToBeClickable(By.xpath("//span[text()='View full profile']")));
                    viewFullProfileButton.click();
                } catch (TimeoutException e) {
                    // If the "View full profile" button is not present, click on the link
                    // corresponding to the name
                    try {
                        String[] parts = name.split(",");
                        name = parts[0];
                        WebElement nameLink = wait.until(ExpectedConditions
                                .elementToBeClickable(By.xpath("//span[text()='" + name + "']/ancestor::a")));
                        nameLink.click();
                    } catch (TimeoutException ex) {
                        System.out.println("No search result found for " + name);
                        Row resultRow = sheet.createRow(rowNum++);
                        resultRow.createCell(0).setCellValue(rowNum - 1);
                        resultRow.createCell(1).setCellValue(name);
                        resultRow.createCell(2).setCellValue(company);
                        resultRow.createCell(3).setCellValue("No results");
                        continue; // Move to the next iteration if the link is not found
                    }
                }

                // Check the current company name on the profile
                try {
                    WebElement companyNameElement = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//span[contains(text(), '" + company + "')]/ancestor::span")));
                    WebElement nextLineElement = null;
                    nextLineElement = companyNameElement.findElement(By.xpath("./following-sibling::*"));
                    String status;
                    if (nextLineElement.getText().toLowerCase().contains("present")) {
                        status = "Same company";
                    } else if (!nextLineElement.getText().toLowerCase().contains("present")) {
                        // Add your specific condition to check for "Recheck"
                        status = "Present keyword is not in right position";
                    } else {
                        status = "Different company";
                    }
                    System.out.println("Status of " + name + " " + status);

                    // Write data to Excel sheet
                    Row row = sheet.createRow(rowNum++);
                    row.createCell(0).setCellValue(rowNum - 1);
                    row.createCell(1).setCellValue(name);
                    row.createCell(2).setCellValue(company);
                    row.createCell(3).setCellValue(status);
                } catch (TimeoutException e) {
                    System.out.println(name + " works in a different company");
                    Row resultRow = sheet.createRow(rowNum++);
                    resultRow.createCell(0).setCellValue(rowNum - 1);
                    resultRow.createCell(1).setCellValue(name);
                    resultRow.createCell(2).setCellValue(company);
                    resultRow.createCell(3).setCellValue("Different company");
                } catch (NoSuchElementException e) {
                    System.out.println("Company details not found in the profile");
                    Row resultRow = sheet.createRow(rowNum++);
                    resultRow.createCell(0).setCellValue(rowNum - 1);
                    resultRow.createCell(1).setCellValue(name);
                    resultRow.createCell(2).setCellValue(company);
                    resultRow.createCell(3).setCellValue("Company details not found in the profile");
                }

                // Go back to the search results page
                driver.navigate().back();
            }
        }finally {
            // Save the workbook in the finally block to ensure it's saved even if an
            // exception occurs
            try (FileOutputStream fileOut = new FileOutputStream("LinkedInProfiles.xlsx")) {
                workbook.write(fileOut);
            } catch (IOException e) {
                e.printStackTrace();
            }

            // Close the browser
            driver.quit();

            long endTime = System.currentTimeMillis();
            long executionTime = endTime - startTime;
            double executionTimeInMinutes = (double) executionTime / 60000.0;
            System.out.println("Execution time: " + executionTimeInMinutes + " minutes");
            System.out.println("Execution time: " + executionTimeInMinutes + " minutes");
        }
    }
}
