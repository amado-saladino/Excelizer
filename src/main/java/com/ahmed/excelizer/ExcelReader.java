package com.ahmed.excelizer;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.*;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.exceptions.InvalidOperationException;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelReader {
	private static XSSFWorkbook workBook;
	private static XSSFSheet sheet;

	private static void setDataSource(String path) {
		File excelFile = new File(path);
		try {
			FileInputStream inputStream = new FileInputStream(excelFile);
			workBook = new XSSFWorkbook(inputStream);
		}
		catch (FileNotFoundException e) {
			e.getMessage();
		}
		catch (IOException e) {
			e.getMessage();
		}
	}

	private static void selectSheet(String sheetName) {
		sheet = workBook.getSheet(sheetName);
	}

	private static String readCell(int rowIndex, int colIndex) {
		return sheet.getRow(rowIndex).getCell(colIndex).getStringCellValue();
	}

	/**
	 * Loads data from Excel file in a tabular form
	 * @param file Excel file
	 * @param sheetName sheet name
	 * @return Tabular form of data set
	 */
	public static Object[][] loadTabularData(String file, String sheetName) {
		setDataSource(file);
		selectSheet(sheetName);

		Object[][] testData = null;

		int rows = sheet.getLastRowNum() + 1;
		int columns = sheet.getRow(0).getLastCellNum();

		testData = new Object[rows - 1][columns];

		for (int i = 1; i < rows; i++) {
			for (int j = 0; j < columns; j++) {
				testData[i - 1][j] = readCell(i, j);
			}
		}
		return testData;
	}

	/**
	 * Iterator over data set from Excel file,
	 * can be used interchangeably with `dataMapIterator` method
	 * @param file Excel file
	 * @param sheetName sheet name
	 * @return Iterator over a set of data
	 */
	public static Iterator<Object[]> dataIterator(String file, String sheetName) {
		setDataSource(file);
		selectSheet(sheetName);
		List<Object[]> data = new ArrayList<>();
		List<String> cells = new ArrayList<>();
		int rows = sheet.getLastRowNum() + 1;
		int columns = sheet.getRow(0).getLastCellNum();

		for (int i = 1; i < rows; i++) {
			cells.clear();
			for (int j = 0; j < columns; j++) {
				cells.add(readCell(i, j));
			}
			data.add(cells.toArray());
		}
		return data.iterator();
	}

	/**
	 * Iterate over data from Excel file
	 * @param file Excel file
	 * @param sheetName sheet name
	 * @return iterator over the data set
	 */
	public static Iterator<Object> dataMapIterator(String file, String sheetName) {
		setDataSource(file);
		selectSheet(sheetName);
		List<Object> data = new ArrayList<>();
		Map row;
		int rows = sheet.getLastRowNum() + 1;
		int columns = sheet.getRow(0).getLastCellNum();

		for (int i = 1; i < rows; i++) {
			row = new HashMap();
			for (int j = 0; j < columns; j++) {
				row.put(readCell(0,j),readCell(i,j));
			}
			data.add(row);
		}
		return data.iterator();
	}

	/**
	 * loads data sheet from Excel file
	 * @param file Excel file name
	 * @param sheetName sheet name in the Excel file
	 * @return data set for the test case
	 */
	public static Object[] getMappedSheet(String file,String sheetName) {
		setDataSource(file);
		selectSheet(sheetName);

		Object[] testData = null;
		Map row;
		int rows = sheet.getLastRowNum() + 1;
		int columns = sheet.getRow(0).getLastCellNum();

		testData = new Object[rows - 1];

		for (int rowIndex = 1; rowIndex < rows; rowIndex++) {
			row = new HashMap();
			for (int columnIndex = 0; columnIndex < columns; columnIndex++) {
				row.put(readCell(0,columnIndex),readCell(rowIndex,columnIndex));
				testData[rowIndex-1] = row;
			}
		}
		return testData;
	}
}
