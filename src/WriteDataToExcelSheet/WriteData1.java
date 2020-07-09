package WriteDataToExcelSheet;

import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WriteData1 {

	public static void main(String[] args) throws IOException {
		// TODO Auto-generated method stub

		FileOutputStream file = new FileOutputStream("D:\\Selenium_Practice\\Testdata1.xlsx");

		XSSFWorkbook workbook = new XSSFWorkbook();
		XSSFSheet sheet1 = workbook.createSheet("Data");
		XSSFSheet sheet2 = workbook.createSheet("Details");

		for (int i = 0; i < 5; i++) {
			XSSFRow Row1 = sheet1.createRow(i);
			XSSFRow Row2 = sheet2.createRow(i);
			for (int j = 0; j < 3; j++) {
				Row1.createCell(j).setCellValue("xyz");
				Row2.createCell(j).setCellValue("ABC");
			}
		}
		workbook.write(file);
		file.close();
		System.out.println("Write Data To Excel Is Done");

	}

}
