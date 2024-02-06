package com.willysalazar.example;

import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;

public class FacebookAutomation {
    public static void main(String[] args) {
        // Set the path to ChromeDriver executable
        System.setProperty("webdriver.chrome.driver", "D:\\ChromeDriver\\119_chromedriver-win64 (1)\\chromedriver-win64\\chromedriver.exe");

        // Initialize ChromeDriver
        WebDriver driver = new ChromeDriver();

        // Set up WebDriverWait
        WebDriverWait wait = new WebDriverWait(driver, 5);

        // Navigate to Facebook login page
        driver.get("https://www.facebook.com/login");

        // Enter email/phone number
        WebElement emailOrPhoneInput = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("email")));
        emailOrPhoneInput.sendKeys("rohithbolla97@gmail.com");

        // Enter password
        WebElement passwordInput = driver.findElement(By.id("pass"));
        passwordInput.sendKeys("Rohith@123");

        // Click on the Login button
        WebElement loginButton = driver.findElement(By.name("login"));
        loginButton.click();

        // Wait for the login process to complete (you may need to adjust the wait condition)
        wait.until(ExpectedConditions.titleContains("Facebook"));


        // Close the browser
        driver.quit();

    }
}
