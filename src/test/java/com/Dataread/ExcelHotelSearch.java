package com.Dataread;

import java.io.*;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.Properties;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.*;
import org.openqa.selenium.chrome.ChromeDriver;
import io.github.bonigarcia.wdm.WebDriverManager;

public class ExcelHotelSearch {
    static WebDriver driver;
    static final String EXCEL_PATH = "src/test/resources/Database_excel/Sample_sheet.xlsx";
    
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

    public static void main(String[] args) {
        WebDriverManager.chromedriver().setup();
        driver = new ChromeDriver();
        driver.get("https://adactinhotelapp.com");

        // Login
        driver.findElement(By.id("username")).sendKeys(getCell(0)); // username
        driver.findElement(By.id("password")).sendKeys(getCell(1)); // password
        driver.findElement(By.id("login")).click();

        // Fill Search Hotel Form
        selectDropdown(By.id("location"), getCell(2));        // Sydney
        selectDropdown(By.id("hotels"), getCell(3));          // Hotel Sunshine
        selectDropdown(By.id("room_type"), getCell(4));       // Double
        selectDropdown(By.id("room_nos"), getCell(5));        // 3 - Three
        clearAndType(By.id("datepick_in"), getCell(6));       // 7/7/2025
        clearAndType(By.id("datepick_out"), getCell(7));      // 8/7/2025
        selectDropdown(By.id("adult_room"), getCell(8));      // 4 - Four
        selectDropdown(By.id("child_room"), getCell(9));      // 1 - One

        // Click Search
        driver.findElement(By.id("Submit")).click();
     // After clicking Search button
        driver.findElement(By.id("radiobutton_0")).click(); // Select first hotel
        driver.findElement(By.id("continue")).click();

        // Fill booking form
        driver.findElement(By.id("first_name")).sendKeys(getCell(10));  // First Name
        driver.findElement(By.id("last_name")).sendKeys(getCell(11));   // Last Name
        driver.findElement(By.id("address")).sendKeys(getCell(12));     // Address
        driver.findElement(By.id("cc_num")).sendKeys(getCell(13));      // Credit Card Number
        driver.findElement(By.id("cc_type")).sendKeys(getCell(14));     // Card Type
        driver.findElement(By.id("cc_exp_month")).sendKeys(getCell(15));// Expiry Month
        driver.findElement(By.id("cc_exp_year")).sendKeys(getCell(16)); // Expiry Year
        driver.findElement(By.id("cc_cvv")).sendKeys(getCell(17));      // CVV

     // Click Book Now
        driver.findElement(By.id("book_now")).click();

        try {
            // Wait for the order number to appear
            Thread.sleep(5000);
        } catch (InterruptedException e) {
            e.printStackTrace();
        }

        // Get Order ID
        WebElement orderIdElement = driver.findElement(By.id("order_no"));
        String orderId = orderIdElement.getAttribute("value");

        // Print in console
        System.out.println("Order ID: " + orderId);

        // Write order ID to Excel (row 18, col 0)
        writeToExcel(18, 0, orderId);
        
    }
    
    public static String getCell(int rowIndex) {
        String value = "";
        try (FileInputStream fis = new FileInputStream(EXCEL_PATH);
             Workbook wb = new XSSFWorkbook(fis)) {
            Sheet sheet = wb.getSheet("Sheet1");
            Row row = sheet.getRow(rowIndex);
            if (row != null) {
                Cell cell = row.getCell(0);
                if (cell != null) {
                    switch (cell.getCellType()) {
                        case STRING: value = cell.getStringCellValue(); break;
                        case NUMERIC:
                            if (DateUtil.isCellDateFormatted(cell)) {
                                value = new SimpleDateFormat("dd/MM/yyyy").format(cell.getDateCellValue());
                            } else {
                                double d = cell.getNumericCellValue();
                                value = (d == (long) d) ? String.valueOf((long) d) : String.valueOf(d);
                            }
                            break;
                        default: value = "";
                    }
                }
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
        return value;
    }
    
    public static void writeToExcel(int rowIndex, int colIndex, String value) {
        FileInputStream fis = null;
        FileOutputStream fos = null;
        Workbook wb = null;
        try {
            File file = new File(EXCEL_PATH);
            fis = new FileInputStream(file);
            wb = new XSSFWorkbook(fis);
            Sheet sheet = wb.getSheet("Sheet1");

            Row row = sheet.getRow(rowIndex);
            if (row == null) {
                row = sheet.createRow(rowIndex);
            }

            Cell cell = row.getCell(colIndex);
            if (cell == null) {
                cell = row.createCell(colIndex);
            }
            cell.setCellValue(value);

            fis.close();  // ✅ Important: Close input stream before writing

            fos = new FileOutputStream(file);
            wb.write(fos);
            System.out.println("Order ID written to Excel successfully at row " + rowIndex);
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            try {
                if (fos != null) fos.close();
                if (wb != null) wb.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }


    public static void selectDropdown(By locator, String value) {
        WebElement element = driver.findElement(locator);
        element.sendKeys(value);
    }

    public static void clearAndType(By locator, String value) {
        WebElement element = driver.findElement(locator);
        element.clear();
        element.sendKeys(value);
    }
    

}
