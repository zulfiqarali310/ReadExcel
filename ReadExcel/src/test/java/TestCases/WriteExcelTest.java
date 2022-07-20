/**
 * 
 */
package TestCases;

import java.io.IOException;

import org.testng.annotations.Test;

import com.writeExcel.WriteExcel;

/**
 * @author PC
 *
 */
public class WriteExcelTest {
	WriteExcel obj = new WriteExcel();
	
	@Test
	public void writeExcelTest() throws Exception {
		
		//obj.write("Test", "male", 0, 2);
		obj.write("Test", "male", 0, 2);
		
	}
	
	@Test
	public void writeExcelTest1() throws Exception {
		
		//obj.write("Test", "male", 0, 2);
		obj.write("Test", "femal", 1, 2);
		
	}

}
