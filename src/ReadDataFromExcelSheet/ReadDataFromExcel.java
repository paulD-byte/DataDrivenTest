package ReadDataFromExcelSheet;

import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadDataFromExcel {

	public static void main(String[] args) throws IOException {
		// DataDrivenTest

		FileInputStream file = new FileInputStream("D:\\Selenium_Practice\\@com.DataDrivenTest\\DataDrivenTest.xlsx");

		XSSFWorkbook workbook = new XSSFWorkbook(file);

		XSSFSheet sheet = workbook.getSheet("Sheet1");// providing sheet name

		// XSSFSheet sheet = workbook.getSheetAt(0);//provide index

		int rowCount = sheet.getLastRowNum();// returns the row count

		int columnCount = sheet.getRow(0).getLastCellNum();// returns cell or column count

		// for getting row and column data

		for (int i = 0; i < rowCount; i++) {
			XSSFRow currentrow = sheet.getRow(i);// focussed on current row

			for (int j = 0; j < columnCount; j++) {
				String value = currentrow.getCell(j).toString();
				System.out.print("  " + value);
			}
			System.out.println();
		}
	}
}
