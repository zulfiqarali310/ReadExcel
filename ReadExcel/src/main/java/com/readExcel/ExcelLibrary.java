package com.readExcel;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelLibrary {

	XSSFWorkbook wb;
	XSSFSheet sheet;

	public ExcelLibrary() throws Exception {

		String excelPath = "C:\\Users\\PC\\eclipse-workspace\\ReadExcel\\TestData\\TestData.xlsx";

		// Create the object of the Class to get excel file.
		File file = new File(excelPath);
		// To Read the file
		FileInputStream fis = new FileInputStream(file);

		// To Create XSSF object represeting XLSX.
		wb = new XSSFWorkbook(fis);

	}

	public String readData(String sheetName, int row, int col) {
		sheet = wb.getSheet(sheetName);

		String data = sheet.getRow(row).getCell(col).getStringCellValue();
		return data;
	}

}
