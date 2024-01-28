package task13;

import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WritreOperation {

	public static void main(String[] args) throws IOException {
		// TODO Auto-generated method stub

	XSSFWorkbook book = new XSSFWorkbook();
	XSSFSheet sheet1 = book.createSheet(); 
	
	Object[][] data = {
		
			{"Name","Age","Email"},
			{"John Doi","30","john@test.com"},
			{"Jane Doe","28","john@test.com"},
			{"Bob Smith","35","jacky@example.com"},
			{"Swapnil","37","swapnil@examole.com"},
	};
	int rowcount = 0;
	for(Object[] row1 : data) {
	XSSFRow row=sheet1.createRow(rowcount++);
	
	int columncount = 0;
	for(Object col : row1) {
		XSSFCell cell =row.createCell(columncount++);
		if(col instanceof String) {
			cell.setCellValue((String)col);
		}else {
			if(col instanceof Integer) {
				cell.setCellValue((int)col);
			}
		}
	try (
	FileOutputStream Output = new FileOutputStream("FirstFile.xlsx");){
	book.write(Output);
	}
		
	}
		
	}
	
	}

}
