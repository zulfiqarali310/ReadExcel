package TestCases;

import org.testng.annotations.Test;

import com.readExcel.ExcelLibrary;

public class ReadExcelTest {

	@Test
	public void readExcelTest() throws Exception {

		ExcelLibrary obj = new ExcelLibrary();
		String strdata = obj.readData("Test", 5, 1);
		System.out.println("The data is: " + strdata);

	}

}
