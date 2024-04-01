package DataDrivenExcel.ExcelDriven;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.NumberToTextConverter;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class dataDriven {

	public ArrayList<String> getData(String testcaseName) throws IOException {
		ArrayList<String> data = new ArrayList<String>();

		FileInputStream fis = new FileInputStream("F:\\Work\\Udemy\\05Dec2023\\Book1.xlsx");
		XSSFWorkbook workbook = new XSSFWorkbook(fis);

		int sheetcount = workbook.getNumberOfSheets();
		for (int i = 0; i < sheetcount; i++) {

			if (workbook.getSheetName(i).equalsIgnoreCase("testdata1")) {

				XSSFSheet sheet = workbook.getSheetAt(i);
				Iterator<Row> rows = sheet.rowIterator();
				Row firstRow = rows.next();

				Iterator<Cell> cell = firstRow.cellIterator();
				int colcount = 0;

				while (cell.hasNext()) {

					Cell value = cell.next();

					if (value.getStringCellValue().equalsIgnoreCase("TestCases")) {

						colcount = value.getColumnIndex();
					}

					// System.out.println(colcount);
				}

				while (rows.hasNext()) {

					Row rowval = rows.next();

					if (rowval.getCell(colcount).getStringCellValue().equalsIgnoreCase(testcaseName)) {
						Iterator<Cell> cellval = rowval.cellIterator();

						while (cellval.hasNext()) {

							Cell val = cellval.next();
							if(val.getCellType()==CellType.STRING) {
								
								data.add(val.getStringCellValue());
							}
							else {	
								
								data.add(NumberToTextConverter.toText(val.getNumericCellValue()));          
							}
							

						}
					}

				}

				
			}

		}
		return data;

	}

}
