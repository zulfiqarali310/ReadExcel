/**
 * 
 */
package com.readExcel;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.annotations.Test;

/**
 * @author PC
 *
 */
public class ReadExcel {
	//@Test
	public void readExcel() throws Exception {

		String excelPath = "C:\\Users\\PC\\eclipse-workspace\\ReadExcel\\TestData\\TestData.xlsx";
		String fileNameString = "TestData";
		String sheetName = "Test";

		// Create the object of the Class to get excel file.
		File file = new File(excelPath);

		// To Read the file
		FileInputStream fis = new FileInputStream(file);

		// To Create XSSF object represeting XLSX.
		XSSFWorkbook wb = new XSSFWorkbook(fis);

		// To Get sheet name
		XSSFSheet sheet = wb.getSheet(sheetName);

		// To Get Row total count
		int totalRowCount = sheet.getLastRowNum();
		System.out.println("Total Rows are: " + totalRowCount);

		// To get Specific row. For Example Get 1st row and 2nd coloumn value.
		String row0AndCol1 = sheet.getRow(0).getCell(1).getStringCellValue();
		System.out.println("Get 1st row and 2nd coloumn value. : " + row0AndCol1);

		// Now we need to print All values column and Row Wis using for Loop
		// Row Print
		for (int i = 0; i <= totalRowCount; i++) {

			Row row = sheet.getRow(i);
			// Get Column of the excel sheet
			for (int j = 0; j < row.getLastCellNum(); j++) {

				String DatarowAndCol = sheet.getRow(i).getCell(j).getStringCellValue();
				System.out.println(DatarowAndCol + " ");

			}
			System.out.println();

		}

		wb.close();

	}

}
