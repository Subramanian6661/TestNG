package com.Dataread;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.Properties;

import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;

import io.github.bonigarcia.wdm.WebDriverManager;

public class Config_value {

    public static WebDriver driver = null;

    public static void main(String[] args) {

        WebDriverManager.chromedriver().setup();
        driver = new ChromeDriver();

        // ✅ Get values from config.properties
        String url = getProperty("url");
        String username = getProperty("username");
        String password = getProperty("password");

        // ✅ Use them
        driver.get(url);
        driver.findElement(By.id("username")).sendKeys(username);
        driver.findElement(By.id("password")).sendKeys(password);
    }

    // ✅ Utility method to read from config.properties
    public static String getProperty(String key) {
        Properties prop = new Properties();
        try {
            FileInputStream fis = new FileInputStream("src/test/resources/config.properties");
            prop.load(fis);
        } catch (IOException e) {
            e.printStackTrace();
        }
        return prop.getProperty(key);
    }
} 
	
	


