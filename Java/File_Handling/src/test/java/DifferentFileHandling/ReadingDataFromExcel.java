package DifferentFileHandling;

import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadingDataFromExcel {

	public static void main(String[] args) throws IOException {
		
		//String filepath = System.getProperty("D:\\1. Kalidas\\2. Codes\\Java\\File_Handling\\TestData\\Dataset_Testing.xlsx");
		
		// Open in Reading Mode
		FileInputStream file =new FileInputStream("D:\\1. Kalidas\\2. Codes\\Java\\File_Handling\\TestData\\Dataset_Testing.xlsx");
		
		// Get control of workbook
		XSSFWorkbook workbook =new XSSFWorkbook(file);
		
		//Go to Particular Sheet
		XSSFSheet Sheet = workbook.getSheet("State_Performance_Report");
		
		
		//Count total number of rows
		int totalRows = Sheet.getLastRowNum();
		
		//Count total number of cells
		 int totalCells = Sheet.getRow(1).getLastCellNum();
		

		//Go to paticular Cell
	
		System.out.println("Number of Row: "+ totalRows );
		System.out.println("Number of Cells: "+ totalCells);
		
		
		for (int r=0;r<totalRows;r++)
		{
			XSSFRow currentRow = Sheet.getRow(r);
			
			for(int c=0;c<totalCells;c++)
			{
				XSSFCell cell = currentRow.getCell(c);
				System.out.print(cell.toString()+"\t");
				
			}
			System.out.println();
		}
		file.close();

	}

}
