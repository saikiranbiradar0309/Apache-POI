package ReadingData;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Demo2 {

	public static void main(String[] args) throws IOException {
		
String excelFilePath = "/Users/saikiranbiradar/Testing/Apache-POI/Excel/Data_Files/Book 6.xlsx";
		
		
		// Opening file in reading mode
		FileInputStream inputStream = new FileInputStream(excelFilePath);
		
		
		// Get the workbook from FileInputStream
		XSSFWorkbook workbook = new XSSFWorkbook(inputStream);
		
		
		// Get the Sheet from Workbook
		XSSFSheet sheet = workbook.getSheet("Sheet1");
		
		
		// Iterator Method
		Iterator iterator = sheet.iterator();
		
		while(iterator.hasNext())
		{
			XSSFRow row = (XSSFRow) iterator.next();
			
			Iterator cellIterator = row.cellIterator();
			
			while(cellIterator.hasNext())
			{
				XSSFCell cell = (XSSFCell) cellIterator.next();
				
				switch(cell.getCellType())
				{
				case STRING: System.out.print(cell.getStringCellValue()); break;
				case NUMERIC: System.out.print(cell.getNumericCellValue()); break;
				case BOOLEAN: System.out.print(cell.getBooleanCellValue()); break;
				}
				
				System.out.print(" | ");
			}
			
			System.out.println();
		}
		
	}
}
