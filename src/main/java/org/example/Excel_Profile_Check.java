package org.example;

import io.github.bonigarcia.wdm.WebDriverManager;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.testng.annotations.BeforeClass;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.time.Duration;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class Excel_Profile_Check {
    WebDriver driver;
    private int rowIndex = 1;
    String path = "C:\\Users\\HEPL\\Downloads\\demo2.xlsx";

    @BeforeClass
    public void setup() {
        WebDriverManager.chromedriver().setup();
        driver = new ChromeDriver();
        driver.manage().window().maximize();
        driver.get("http://216.48.191.170/budgie_test/public/index.php");
        driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(10));
        driver.manage().timeouts().pageLoadTimeout(Duration.ofSeconds(10));
    }

    @Test(dataProvider = "dataProvider",priority = 1)
    public void profileCheck(String user, String pass, String status) throws InterruptedException {
        driver.findElement(By.id("employee_id")).sendKeys(user);
        driver.findElement(By.id("login_password")).sendKeys(pass);
        driver.findElement(By.id("btnLogin")).click();
        WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(10));
        wait.until(ExpectedConditions.visibilityOfElementLocated(By.partialLinkText("Home")));
        driver.navigate().to("http://216.48.191.170/budgie_test/public/index.php/candidate_profile");
        wait.until(ExpectedConditions.visibilityOfElementLocated(By.partialLinkText("Home")));
        List<WebElement> list = driver.findElements(By.className("row"));
        boolean isNullFound = false;
        for (WebElement out : list) {
            String find = out.getAttribute("textContent").trim();
            Pattern pattern = Pattern.compile("\\s");
            Matcher matcher = pattern.matcher(find);
            String result = matcher.replaceAll("");

            if (result.matches("(?i).*\\bnull\\b.*")) {
                isNullFound = true;
                break; // Exit the loop once "null" is found
            }
        }
        if (isNullFound) {
            System.out.println("Null");
            updateExcel(rowIndex,"Null");
        } else {
            updateExcel(rowIndex,"NO");
            System.out.println("NO");
        }
        rowIndex++;
        driver.navigate().to("http://216.48.191.170/budgie_test/public/index.php/logout");
    }

    public void updateExcel(int rowIndex,String name) {
        try (FileInputStream input = new FileInputStream(path);
             XSSFWorkbook workbook = new XSSFWorkbook(input);
             FileOutputStream out = new FileOutputStream(path)) {
            XSSFSheet sheet = workbook.getSheetAt(0);
            XSSFRow row = sheet.getRow(rowIndex);
            row.getCell(2).setCellValue(name);
            workbook.write(out);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
    @DataProvider
    public Object[][] dataProvider() throws IOException {
        FileInputStream input = new FileInputStream(path);
        XSSFWorkbook workbook = new XSSFWorkbook(input);
        XSSFSheet sheet = workbook.getSheetAt(0);
        int rowNum = sheet.getLastRowNum();
        int col = sheet.getRow(1).getLastCellNum();
        String dataProvider[][] = new String[rowNum][col];
        for (int i = 1; i <= rowNum; i++) {
            XSSFRow rows = sheet.getRow(i);
            for (int j = 0; j < col; j++) {
                XSSFCell cells = rows.getCell(j);
                DataFormatter format = new DataFormatter();
                String data = format.formatCellValue(cells);
                dataProvider[i - 1][j] = data;
            }
        }
        workbook.close();
        return dataProvider;
    }
}