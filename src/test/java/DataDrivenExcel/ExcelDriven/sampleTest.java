package DataDrivenExcel.ExcelDriven;

import java.io.IOException;
import java.util.ArrayList;

public class sampleTest {

	public static void main(String[] args) throws IOException {
		// TODO Auto-generated method stub
		
		dataDriven dd = new dataDriven();
		
		ArrayList<String> exceldata = dd.getData("purchase");

		System.out.println(exceldata.get(0));
		System.out.println(exceldata.get(1));
		System.out.println(exceldata.get(2));
		System.out.println(exceldata.get(3));
	}

}
