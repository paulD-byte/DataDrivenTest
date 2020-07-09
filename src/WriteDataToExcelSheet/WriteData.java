package WriteDataToExcelSheet;

import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WriteData {

	public static void main(String[] args) throws IOException {
		// Write Data To Excel Sheet

		FileOutputStream file = new FileOutputStream("D:\\Selenium_Practice\\Testdata.xlsx");

		XSSFWorkbook workbook = new XSSFWorkbook();
		XSSFSheet sheet = workbook.createSheet("Data");

		for (int i = 0; i < 5; i++) {
			XSSFRow createRow = sheet.createRow(i);
			for (int j = 0; j < 3; j++) {
				createRow.createCell(j).setCellValue("xyz");
			}
		}
		workbook.write(file);
		file.close();
		System.out.println("Write Data To Excel Is Done");

	}

}
