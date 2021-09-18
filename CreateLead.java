package Week5.day2;

import java.io.IOException;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.testng.annotations.Test;
import org.testng.annotations.DataProvider;

public class CreateLead extends BaseClass {

	public XSSFWorkbook wb;
	public XSSFSheet ws;

	@Test(dataProvider = "sendData")
	public void runCreateLead(String company, String fName, String lName, String ph) {
		driver.findElement(By.linkText("Create Lead")).click();
		driver.findElement(By.id("createLeadForm_companyName")).sendKeys(company);
		driver.findElement(By.id("createLeadForm_firstName")).sendKeys(fName);
		driver.findElement(By.id("createLeadForm_lastName")).sendKeys(lName);
		driver.findElement(By.id("createLeadForm_primaryPhoneNumber")).sendKeys(ph);
		driver.findElement(By.name("submitButton")).click();
	}

	@DataProvider
	public String[][] sendData() throws IOException {
		String[][] data;
		wb = new XSSFWorkbook("./data/LeaftapsInputFile.xlsx");
		ws = wb.getSheet("CreateLead");
		int rowCount = ws.getLastRowNum();
		int cellCount = ws.getRow(0).getLastCellNum();
		System.out.println("Rows = " + rowCount + " and Columns = " + cellCount);
		data = new String[rowCount][cellCount];
		for (int i = 1; i <= rowCount; i++) {
			for (int j = 0; j < cellCount; j++) {
				String text = ws.getRow(i).getCell(j).getStringCellValue();
				System.out.println(text);
				data[i - 1][j] = text;
				System.out.println(data[i - 1][j]);
			}

		}

		wb.close();
		return data;

	}

}
