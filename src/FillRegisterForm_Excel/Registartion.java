package FillRegisterForm_Excel;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.concurrent.TimeUnit;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.support.ui.Select;

public class Registartion {

	public static void main(String[] args) throws IOException {
		// Filling A Registration Form

		// Setup

		WebDriver driver = new FirefoxDriver();
		driver.get("http://newtours.demoaut.com/");

		driver.manage().window().maximize();
		driver.manage().deleteAllCookies();

		driver.manage().timeouts().pageLoadTimeout(30, TimeUnit.SECONDS);
		driver.manage().timeouts().implicitlyWait(20, TimeUnit.SECONDS);

		// Getting Data From Excel Sheet

		FileInputStream file = new FileInputStream("D:\\Selenium_Practice\\@com.DataDrivenTest\\Registration.xlsx");

		XSSFWorkbook workbook = new XSSFWorkbook(file);
		XSSFSheet sheet = workbook.getSheet("Details");// providing sheet name

		// XSSFSheet sheet = workbook.getSheetAt(0);//provide index

		int rowCount = sheet.getLastRowNum();// returns the row count
		System.out.println("No. of Rows:" + rowCount);

		// for getting row data

		for (int row = 1; row < rowCount; row++) {
			XSSFRow currentrow = sheet.getRow(row);// focussed on current row

			String fn = currentrow.getCell(0).toString();
			String ln = currentrow.getCell(1).toString();
			String phn = currentrow.getCell(2).toString();
			String email = currentrow.getCell(3).toString();
			String address = currentrow.getCell(4).toString();
			String city = currentrow.getCell(5).toString();
			String state = currentrow.getCell(6).toString();
			String pincode = currentrow.getCell(7).toString();
			String country = currentrow.getCell(8).toString();
			String uname = currentrow.getCell(9).toString();
			String pwd = currentrow.getCell(10).toString();

			// Registration Process

			driver.findElement(By.linkText("REGISTER")).click();

			// Entering Contact Details

			driver.findElement(By.name("firstName")).sendKeys(fn);
			driver.findElement(By.name("lastName")).sendKeys(ln);
			driver.findElement(By.name("phone")).sendKeys(phn);
			driver.findElement(By.id("userName")).sendKeys(email);

			// Entering Mailing Details

			driver.findElement(By.name("address1")).sendKeys(address);
			driver.findElement(By.name("city")).sendKeys(city);
			driver.findElement(By.name("state")).sendKeys(state);
			driver.findElement(By.name("postalCode")).sendKeys(pincode);

			// Country Selection

			Select select = new Select(driver.findElement(By.name("country")));
			select.selectByVisibleText(country);

			// Entering User Information

			driver.findElement(By.id("email")).sendKeys(uname);
			driver.findElement(By.name("password")).sendKeys(pwd);
			driver.findElement(By.name("confirmPassword")).sendKeys(pwd);

			// Submitting Form
			driver.findElement(By.name("register")).click();

			// Validation

			WebElement text = driver.findElement(By.xpath("//font[contains(text(),'Thank you for registering.')]"));
			System.out.println(text);

			if (text.equals(
					driver.findElement(By.xpath("//font[contains(text(),'Thank you for registering.')]"))) == true) {
				System.out.println("Test Passed");
			} else {
				System.out.println("Test Failed");

			}
		}
		driver.close();
	}

}
