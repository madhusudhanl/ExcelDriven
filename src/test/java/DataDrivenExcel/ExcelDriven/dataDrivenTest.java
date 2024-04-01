package DataDrivenExcel.ExcelDriven;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;

public class dataDrivenTest {
	
	DataFormatter formatter = new DataFormatter();
	
	@Test(dataProvider = "exceldata")
	public void testCaseData(String testcaseName,String name,String address,String id) {
		
		System.out.println(testcaseName+"--"+name+"--"+	address+"--"+id);
	}
	
	@DataProvider(name="exceldata")
	public Object[][] getData() throws IOException {
		
//		Object[][] data = {
//							{"Name1","Address1",123},
//							{"Name2","Address2",456},
//							{"Name3","Address3",789}
//							
//							};
//		return data;
		
		FileInputStream fis = new FileInputStream("F:\\Work\\Udemy\\05Dec2023\\Book1.xlsx");
		XSSFWorkbook workbook = new XSSFWorkbook(fis);
		XSSFSheet sheet = workbook.getSheetAt(0);
		int rowCount = sheet.getPhysicalNumberOfRows();
		XSSFRow row = sheet.getRow(0);
		short colcount = row.getLastCellNum();
		
		Object data[][] = new Object[rowCount-1][colcount]; // rowcount-1 because we dont need header.
		for(int i=0;i<rowCount-1;i++) {
			
			row=sheet.getRow(i+1);  // because i=0.
			for(int j=0 ; j<colcount; j++) {
				
				XSSFCell cell = row.getCell(j);
				//formatter.formatCellValue(cell); used to convert to string value.
				
				data[i][j] = formatter.formatCellValue(cell);
			}
			
		}
		return data;
		
	}

}
