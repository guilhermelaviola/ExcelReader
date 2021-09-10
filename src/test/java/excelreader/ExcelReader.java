package excelreader;

import java.io.IOException;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelReader {

	// By making these variables static, that means they can be accessed globally
	static XSSFWorkbook workbook;
	static XSSFSheet sheet;

	// Constructor
	public ExcelReader (String excelPath, String sheetName) {

		try {			
			// Creating the workbook and sheet objects
			workbook = new XSSFWorkbook(excelPath);
			sheet = workbook.getSheet(sheetName);

			// Closing the workbook object
			workbook.close();
		
		} catch (Exception e) {
			// Catching cause and message
			System.out.println(e.getCause());
			System.out.println(e.getMessage());
			e.printStackTrace();
		}
	}

	public void getCellData(int rowNum, int colNum) throws IOException {
		// Creating the DataFormatter object
		DataFormatter formatter = new DataFormatter();

		// Storing the row and column numbers into the value variable and printing it
		Object value = formatter.formatCellValue(sheet.getRow(rowNum).getCell(colNum));
		System.out.println("Data requested: " + value);

	}

	public void getRowCount() {
		// Variable to store the number of rows and printing it
		int rowCount = sheet.getPhysicalNumberOfRows();
		System.out.println("Number of rows: " + rowCount);
	}
}