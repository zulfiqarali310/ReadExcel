package com.writeExcel;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WriteExcel {

	XSSFWorkbook wb;
	XSSFSheet sheet;

	public void write(String sheetName, String cellvalue, int row, int col) throws IOException {

		String excelPath = "C:\\Users\\PC\\eclipse-workspace\\ReadExcel\\TestData\\TestData.xlsx";

		// Create the object of the Class to get excel file.
		File file = new File(excelPath);
		// To Read the file
		FileInputStream fis = new FileInputStream(file);

		// To Create XSSF object represeting XLSX.
		wb = new XSSFWorkbook(fis);

		// To Get sheet name
		XSSFSheet sheet = wb.getSheet(sheetName);
		
		sheet.getRow(row).createCell(col).setCellValue(cellvalue);
		
		FileOutputStream fos = new FileOutputStream(new File(excelPath));
		wb.write(fos);
		fos.close();

	}

}
