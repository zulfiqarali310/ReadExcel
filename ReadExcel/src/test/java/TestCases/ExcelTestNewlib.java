package TestCases;

import org.testng.annotations.Test;

import com.utility.NewExcelLibrary;

public class ExcelTestNewlib {

	NewExcelLibrary obj = new NewExcelLibrary("C:\\Users\\PC\\eclipse-workspace\\ReadExcel\\TestData\\TestData.xlsx");

	@Test(priority = 1)
	public void TestCase1() {

		int totalrows = obj.getRowCount("Employee");
		System.out.println("Total rows: " + totalrows);
	}

	@Test(priority = 2)
	public void TestCase2() {

		String salary = obj.getCellData("Employee", "Salary", 4);
		System.out.println("Total rows: " + salary);
	}

	@Test(priority = 3)
	public void TestCase3() {

		String rating = obj.getCellData("Employee", 4, 2);
		System.out.println("Total rows: " + rating);

		int a = 1;
		Double sum = Double.valueOf(rating) + a;
		System.out.println("The rating is now " + sum);
	}

	@Test(priority = 4)
	public void TestCase4() {

		boolean setData = obj.setCellData("Employee", "ID", 2, "105");
		System.out.println("Record updated: " + setData);

	}

	@Test(priority = 5)
	public void TestCase5() {

		boolean flag = obj.addSheet("NewSheet");
		System.out.println("New sheet added: " + flag);

	}

	@Test(priority = 6)
	public void TestCase6() {

		boolean sheet1 = obj.removeSheet("NewSheet");
		System.out.println("The sheet is removed: " + sheet1);

	}

}
