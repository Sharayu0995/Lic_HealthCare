package com.practice.LicHealth;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.Calendar;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.Keys;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;

public class Xls_Reader 
{
	private String path;

	 FileInputStream fis = null;
	 FileOutputStream fileOut = null;
	 XSSFWorkbook workbook = null;
	 XSSFSheet sheet = null;
	 XSSFRow row = null;
	 XSSFCell cell = null;

	public Xls_Reader (String path) 
	{

		this.path = path;
		try {
			fis = new FileInputStream(path);
			workbook = new XSSFWorkbook(fis);
			sheet = workbook.getSheetAt(0);
			fis.close();
			
		} catch (Exception e)
		{
			
			e.printStackTrace();
		}
	}
	public void setCellData(String sheetName, String colName, int rowNum, String data)
	{
		try {
			fis = new FileInputStream(path);
			workbook = new XSSFWorkbook(fis);

			if (rowNum <= 0)
				return ;

			int index = workbook.getSheetIndex(sheetName);
			int colNum = -1;
			if (index == -1)
				return ;

			sheet = workbook.getSheetAt(index);

			row = sheet.getRow(0);
			for (int i = 0; i < row.getLastCellNum(); i++) {
				// System.out.println(row.getCell(i).getStringCellValue().trim());
				if (row.getCell(i).getStringCellValue().trim().equals(colName))
					colNum = i;
			}
			if (colNum == -1)
				return ;

			sheet.autoSizeColumn(colNum);
			row = sheet.getRow(rowNum - 1);
			if (row == null)
				row = sheet.createRow(rowNum - 1);

			cell = row.getCell(colNum);
			if (cell == null)
				cell = row.createCell(colNum);

			// cell style
			// CellStyle cs = workbook.createCellStyle();
			// cs.setWrapText(true);
			// cell.setCellStyle(cs);
			cell.setCellValue(data);

			fileOut = new FileOutputStream(path);

			workbook.write(fileOut);

			fileOut.close();

		} catch (Exception e) {
			e.printStackTrace();
			return;
		}
		return;
	}
	public void addSheet(String sheetname)
	{

		FileOutputStream fileOut;
		try {
			workbook.createSheet(sheetname);
			fileOut = new FileOutputStream(path);
			workbook.write(fileOut);
			fileOut.close();
		} catch (Exception e) {
			e.printStackTrace();
			return ;
		}
		return ;
	}
	public String getCellData(String sheetName, String colName, int rowNum) {
		try {
			if (rowNum <= 0)
				return "";

			int index = workbook.getSheetIndex(sheetName);
			int col_Num = -1;
			if (index == -1)
				return "";

			sheet = workbook.getSheetAt(index);
			row = sheet.getRow(0);
			for (int i = 0; i < row.getLastCellNum(); i++) {
				// System.out.println(row.getCell(i).getStringCellValue().trim());
				if (row.getCell(i).getStringCellValue().trim().equals(colName.trim()))
					col_Num = i;
			}
			if (col_Num == -1)
				return "";

			sheet = workbook.getSheetAt(index);
			row = sheet.getRow(rowNum - 1);
			if (row == null)
				return "";
			cell = row.getCell(col_Num);

			if (cell == null)
				return "";
			// System.out.println(cell.getCellType());
			if (cell.getCellType() == Cell.CELL_TYPE_STRING)
				return cell.getStringCellValue();
			else if (cell.getCellType() == Cell.CELL_TYPE_NUMERIC || cell.getCellType() == Cell.CELL_TYPE_FORMULA) {

				String cellText = String.valueOf(cell.getNumericCellValue());
				if (HSSFDateUtil.isCellDateFormatted(cell)) {
					// format in form of M/D/YY
					double d = cell.getNumericCellValue();

					Calendar cal = Calendar.getInstance();
					cal.setTime(HSSFDateUtil.getJavaDate(d));
					cellText = (String.valueOf(cal.get(Calendar.YEAR))).substring(2);
					cellText = cal.get(Calendar.DAY_OF_MONTH) + "/" + cal.get(Calendar.MONTH) + 1 + "/" + cellText;

					// System.out.println(cellText);

				}

				return cellText;
			} else if (cell.getCellType() == Cell.CELL_TYPE_BLANK)
				return "";
			else
				return String.valueOf(cell.getBooleanCellValue());

		} catch (Exception e) {

			e.printStackTrace();
			return "row " + rowNum + " or column " + colName + " does not exist in xls";
		}
	}

	// returns the data from a cell
 	public String getCellData(String sheetName, int colNum, int rowNum)
	{
		try {
			if (rowNum <= 0)
				return "";

			int index = workbook.getSheetIndex(sheetName);

			if (index == -1)
				return "";

			sheet = workbook.getSheetAt(index);
			row = sheet.getRow(rowNum - 1);
			if (row == null)
				return "";
			cell = row.getCell(colNum);
			if (cell == null)
				return "";

			if (cell.getCellType() == Cell.CELL_TYPE_STRING)
				return cell.getStringCellValue();
			else if (cell.getCellType() == Cell.CELL_TYPE_NUMERIC || cell.getCellType() == Cell.CELL_TYPE_FORMULA) {

				String cellText = String.valueOf(cell.getNumericCellValue());
				if (HSSFDateUtil.isCellDateFormatted(cell)) {
					// format in form of M/D/YY
					double d = cell.getNumericCellValue();

					Calendar cal = Calendar.getInstance();
					cal.setTime(HSSFDateUtil.getJavaDate(d));
					cellText = (String.valueOf(cal.get(Calendar.YEAR))).substring(2);
					cellText = cal.get(Calendar.MONTH) + 1 + "/" + cal.get(Calendar.DAY_OF_MONTH) + "/" + cellText;

					// System.out.println(cellText);

				}

				return cellText;
			} else if (cell.getCellType() == Cell.CELL_TYPE_BLANK)
				return "";
			else
				return String.valueOf(cell.getBooleanCellValue());
		} catch (Exception e) {

			e.printStackTrace();
			return "row " + rowNum + " or column " + colNum + " does not exist  in xls";
		}
	}
	
	public int getRowCount(String sheetName) {
		int index = workbook.getSheetIndex(sheetName);
		if (index == -1)
			return 0;
		else {
			sheet = workbook.getSheetAt(index);
			int number = sheet.getLastRowNum() + 1;
			return number;
		}

	}
	public boolean addColumn(String sheetName, String colName) {
		// System.out.println("**************addColumn*********************");

		try {
			fis = new FileInputStream(path);
			workbook = new XSSFWorkbook(fis);
			int index = workbook.getSheetIndex(sheetName);
			if (index == -1)
				return false;

			XSSFCellStyle style = workbook.createCellStyle();
			style.setFillForegroundColor(HSSFColor.GREY_40_PERCENT.index);
			style.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);

			sheet = workbook.getSheetAt(index);

			row = sheet.getRow(0);
			if (row == null)
				row = sheet.createRow(0);

			// cell = row.getCell();
			// if (cell == null)
			// System.out.println(row.getLastCellNum());
			if (row.getLastCellNum() == -1)
				cell = row.createCell(0);
			else
				cell = row.createCell(row.getLastCellNum());

			cell.setCellValue(colName);
			cell.setCellStyle(style);

			fileOut = new FileOutputStream(path);
			workbook.write(fileOut);
			fileOut.close();

		} catch (Exception e) {
			e.printStackTrace();
			return false;
		}

		return true;

	}
	protected static void  newbranch()
	{
		//search for another code
		System.setProperty("webdriver.chrome.driver", "C:\\Users\\fidel\\eclipse-workspace\\LicHealth\\Driver\\chromedriver.exe");
		WebDriver driver1=new ChromeDriver();
		driver1.manage().window().maximize();
		driver1.get("https://www.icicilombard.com/campaigns/health-insurance/health-insurance-mumbai");
		System.out.println(driver1.getTitle());
		driver1.findElement(By.cssSelector("body")).sendKeys(Keys.CONTROL +"\t");
			
			JavascriptExecutor jse1 = (JavascriptExecutor)driver1;
    		jse1.executeScript("window.scrollBy(0,950)");
    		
    		
    		 // List of hospital
            List<WebElement> rows1=driver1.findElements(By.xpath("//table[@id='HospitalList']//tr"));
          	System.out.println("row size="+(rows1.size()-1));
          	int rowCount1=rows1.size();
          	
          	
          	
         //scroll the screen
          		JavascriptExecutor jse11 = (JavascriptExecutor)driver1;
          		jse11.executeScript("window.scrollBy(0,950)");
          		//Thread.sleep(10000);
          		driver1.findElement(By.id("Campaign1_C019_LnkSearch")).click();
          		//elementCSV.click();
          		
          			//For hospital
          		String beforexpath_hos1="//*[@id='HospitalList']//tr[";
          		String afterxpath_hos1="]/td[2]";
          	
          		
          		//Address
          		String beforexpath_Address1="//*[@id='HospitalList']//tr[";
          		String afterxpath_Address1="]/td[3]";


          		//City 
          		String beforexpath_City1="//*[@id='HospitalList']//tr[";
          		String afterxpath_City1="]/td[4]";
          		
          	
          		
          		//State
          		String beforexpath_state1="//*[@id='HospitalList']//tr[";
          		String afterxpath_state1="]/td[5]";
          		
          		//Contact
          		String beforexpath_contact1="//*[@id='HospitalList']//tr[";
          		String afterxpath_contact1="]/td[6]";
          		
          		
          		Xls_Reader reader=new Xls_Reader("C:\\Users\\fidel\\eclipse-workspace\\LicHealth\\Driver\\Book1.xlsx");
              	reader.addSheet("TestData2");
              	
              	reader.addColumn("TestData2","Hospital Name");
              	reader.addColumn("TestData2","Address");
              	reader.addColumn("TestData2","City");
              	reader.addColumn("TestData2","State");
              	reader.addColumn("TestData2","Contact");
          			
          				for(int y=2;y<=rowCount1-1;y++)
          				{
          				//Hospital name
          				String actualxpath_hos1=beforexpath_hos1+y+afterxpath_hos1;

          				String hosame1= driver1.findElement(By.xpath(actualxpath_hos1)).getText();
          				System.out.println(hosame1);
          				
          				reader.setCellData("TestData2","Hospital Name",y, hosame1);
          				//reader.getCellData("TestData", "Hospital Name", i);
          				
          				//Address
          				String actualxpath_add1=beforexpath_Address1+y+afterxpath_Address1;
          				String Address= driver1.findElement(By.xpath(actualxpath_add1)).getText();
          				System.out.println(Address);
          				reader.setCellData("TestData2","Address",y, Address);
          				//reader.getCellData("TestData", "Address", i);
          				
          				//City
          				String actualxpath_city1=beforexpath_City1+y+afterxpath_City1;
          				String City= driver1.findElement(By.xpath(actualxpath_city1)).getText();
          				System.out.println(City);		
          				reader.setCellData("TestData2","City",y, City);
          				//reader.getCellData("TestData", "City", i);
          				
          				//State
          				String actualxpath_state1=beforexpath_state1+y+afterxpath_state1;
          				String state= driver1.findElement(By.xpath(actualxpath_state1)).getText();
          				System.out.println(state);
          				reader.setCellData("TestData2","State",y, state);
          				
          				//reader.getCellData("TestData", "State", i);
          				//CONTACTNO
          				String actualxpath_cont1=beforexpath_contact1+y+afterxpath_contact1;
          				String contact= driver1.findElement(By.xpath(actualxpath_cont1)).getText();
          				System.out.println(contact);
          				reader.setCellData("TestData2","Contact",y, contact);
          				
          				}
	}	

	
	
}
