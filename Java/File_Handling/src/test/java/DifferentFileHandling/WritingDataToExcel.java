package DifferentFileHandling;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WritingDataToExcel {

	public static void main(String[] args) throws FileNotFoundException {
		
		XSSFWorkbook workbook =new XSSFWorkbook();
	
		FileOutputStream file = new FileOutputStream("D:\\1. Kalidas\\2. Codes\\Java\\File_Handling\\TestData\\NewWriteExcel.xlsx");
		
		XSSFSheet sheet = workbook.createSheet("Imp_Data");
		

	}

}
