package Week5.day2;

import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadExcel {

	
	public static void main(String[] args) throws IOException {
		XSSFWorkbook wb = new XSSFWorkbook("./data/CreateLead.xlsx");
		XSSFSheet ws=wb.getSheet("Sheet1");
		int rowcount=ws.getLastRowNum();
		int Cellcount=ws.getRow(0).getLastCellNum();
		for(int i=1;i<=rowcount;i++) {
			for(int j=0;j<Cellcount;j++) {
				String text=ws.getRow(i).getCell(j).getStringCellValue();
				System.out.println(text);
			}
		}
		wb.close();
	}
}
