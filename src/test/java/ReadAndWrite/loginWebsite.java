package ReadAndWrite;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;

import java.io.FileInputStream;

public class loginWebsite {
    public static XSSFSheet excelSheet;
    public static XSSFCell cell;
    public static WebDriver driver;
    public static XSSFWorkbook ExcelWBook;
    public static XSSFSheet ExcelWSheet;

    public static void main(String[] args) throws InterruptedException {
        System.setProperty("webdriver.gecko.driver", System.getProperty("user.dir") + "/src/main/resources/geckodriver");
        driver = new ChromeDriver();


        excelSheet = loginWebsite.readExcel("/Users/hakanozcan/Desktop/loginData.xlsx", "TestSheet");
        for (int i = 0; i <= 1; i++) {

            driver.get("https://www.facebook.com/");

            driver.findElement(By.id("name"))
                    .sendKeys(excelSheet.getRow(i).getCell(0).getStringCellValue());
            driver.findElement(By.id("pass"))
                    .sendKeys(excelSheet.getRow(i).getCell(1).getStringCellValue());

            driver.findElement(By.id("u_0_b")).click();
        }
        driver.close();
    }

    public static XSSFSheet readExcel(String Path, String SheetName) {
        try {
            System.out.println(Path);

            FileInputStream ExcelFile = new FileInputStream(Path);

            ExcelWBook = new XSSFWorkbook(ExcelFile);
            ExcelWSheet = ExcelWBook.getSheet(SheetName);
        } catch (Exception e) {
            System.out.println(e);
        }
        return ExcelWSheet;
    }
}



