package com.ahmed.excelizer;

import java.io.File;
import java.io.IOException;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.exceptions.InvalidOperationException;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelReader {
	private static XSSFWorkbook workBook;
	private static XSSFSheet sheet;

	private static void setDataSource(String path) throws InvalidFormatException, IOException {
		// try {
		workBook = new XSSFWorkbook(new File(path));
		// } /*catch (Throwable e) {
		// System.err.println("Error: Excel workbook cannot be found! Please check file
		// path");
		// System.err.println( e.getMessage() );
		// }*/
	}

	private static void selectSheet(String sheetName) {
		sheet = workBook.getSheet(sheetName);
	}

	private static String readCell(int rowIndex, int colIndex) {
		return sheet.getRow(rowIndex).getCell(colIndex).getStringCellValue();
	}

	/**
	 * Creates a data source from a file and sheet name
	 * 
	 * @param file      path to Excel file
	 * @param sheetName sheet name inside the excel workbook
	 * @return Array to use as a data source
	 */
	public static Object[][] loadTestData(String file, String sheetName) {
		Object[][] testData = null;
		try {
			setDataSource(file);
			selectSheet(sheetName);
			int rows = sheet.getLastRowNum() + 1;
			int columns = sheet.getRow(0).getLastCellNum();
			testData = new Object[rows - 1][columns];

			for (int i = 1; i < rows; i++) {
				for (int j = 0; j < columns; j++) {
					testData[i - 1][j] = readCell(i, j);
				}
			}
		} catch (InvalidFormatException | InvalidOperationException | IOException e) {
			System.err.println("Error: Excel workbook cannot be found! Please check file path");
			System.err.println(e.getMessage());
		} catch (NullPointerException e) {
			System.err.println("Error: Cannot select the given sheet '" + sheetName + ".' Please check the sheet name");			
		}		
		return testData;
	}
}
