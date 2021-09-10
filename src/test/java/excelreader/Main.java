package excelreader;

import java.io.IOException;
import java.util.Scanner;

// Testing the program for Excel data manipulation 
public class Main {
	
	public static void main(String[] args) throws IOException {
		
		// Creating Scanner object
		Scanner s = new Scanner(System.in);
		
		// Indicating where the file path is located in the system and the sheet name 
		String excelPath = "./data/CustomerTest.xlsx";
		String sheetName = "Sheet1";
		
		// Creating the ExcelUtils object
		ExcelReader excel = new ExcelReader(excelPath, sheetName);
		
		System.out.println("GETTING CELL DATA:");
		// Invoking getRowCount()
		excel.getRowCount();
		
		System.out.println("GETTING CELL DATA:");
		// Variables to store data entered by user through Scanner
		System.out.print("Row index: ");
		int a = s.nextInt();
		System.out.print("Column index: ");
		int b = s.nextInt();
		
		// Invoking getCellData() with user entered data
		excel.getCellData(a, b);;
		
		// Closing Scanner object
		s.close();
	}
}
