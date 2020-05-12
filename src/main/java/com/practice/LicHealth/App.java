package com.practice.LicHealth;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.FileWriter;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.dom4j.io.XPPReader;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.Keys;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;


class App extends Xls_Reader
{
  
	

	public App(String path) {
		super(path);
		System.out.println("HI");
	}

	public static void main( String[] args ) throws Exception
    {
    	 
        System.out.println( "Hello World!" );
        System.setProperty("webdriver.chrome.driver", "C:\\Users\\fidel\\eclipse-workspace\\LicHealth\\Driver\\chromedriver.exe");
        WebDriver driver=new ChromeDriver();
        driver.manage().window().maximize();
        driver.get("https://www.icicilombard.com/campaigns/health-insurance/health-insurance-pune");
        //another city
        
        
        List<WebElement> rows=driver.findElements(By.xpath("//table[@id='HospitalList']//tr"));
      	System.out.println("row size="+(rows.size()-1));
      	int rowCount=rows.size();
      	
      	
      	
      //scroll the screen
      		JavascriptExecutor jse = (JavascriptExecutor)driver;
      		jse.executeScript("window.scrollBy(0,950)");
      		//Thread.sleep(10000);
      		driver.findElement(By.id("Campaign1_C017_LnkSearch")).click();
      		//elementCSV.click();
      		
      		//For hospital
      		String beforexpath_hos="//*[@id='HospitalList']/tbody/tr[";
      		String afterxpath_hos="]/td[2]";
      		
      		
      		//Address
      		String beforexpath_Address="//*[@id='HospitalList']/tbody/tr[";
      		String afterxpath_Address="]/td[3]/span";


      		//City 
      		String beforexpath_City="//*[@id='HospitalList']/tbody/tr[";
      		String afterxpath_City="]/td[4]/span";
      		
      	
      		
      		//State
      		String beforexpath_state="//*[@id='HospitalList']/tbody/tr[";
      		String afterxpath_state="]/td[5]/span";
      		
      		//Contact
      		String beforexpath_contact="//*[@id='HospitalList']/tbody/tr[";
      		String afterxpath_contact="]/td[6]/span";
      		
      		
      		Xls_Reader reader=new Xls_Reader("C:\\Users\\fidel\\eclipse-workspace\\LicHealth\\Driver\\Book1.xlsx");
          	reader.addSheet("TestData");
          	
          	reader.addColumn("TestData","Company Name");
          	reader.addColumn("TestData","Address");
          	reader.addColumn("TestData","City");
          	reader.addColumn("TestData","State");
          	reader.addColumn("TestData","Contact");
      			
      				for(int i=2;i<=rowCount;i++)
      				{
      				//Hospital name
      				String actualxpath_hos=beforexpath_hos+i+afterxpath_hos;

      				String hosame= driver.findElement(By.xpath(actualxpath_hos)).getText();
      				System.out.println(hosame);
      				
      				reader.setCellData("TestData","Company Name",i, hosame);
      				//reader.getCellData("TestData", "Company Name", i);
      				
      				//Address
      				String actualxpath_add=beforexpath_Address+i+afterxpath_Address;
      				String Address= driver.findElement(By.xpath(actualxpath_add)).getText();
      				System.out.println(Address);
      				reader.setCellData("TestData","Address",i, Address);
      				//reader.getCellData("TestData", "Address", i);
      				
      				//City
      				String actualxpath_city=beforexpath_City+i+afterxpath_City;
      				String City= driver.findElement(By.xpath(actualxpath_city)).getText();
      				System.out.println(City);		
      				reader.setCellData("TestData","City",i, City);
      				//reader.getCellData("TestData", "City", i);
      				
      				//State
      				String actualxpath_state=beforexpath_state+i+afterxpath_state;
      				String state= driver.findElement(By.xpath(actualxpath_state)).getText();
      				System.out.println(state);
      				reader.setCellData("TestData","State",i, state);
      				
      				//reader.getCellData("TestData", "State", i);
      				//CONTACTNO
      				String actualxpath_cont=beforexpath_contact+i+afterxpath_contact;
      				String contact= driver.findElement(By.xpath(actualxpath_cont)).getText();
      				System.out.println(contact);
      				reader.setCellData("TestData","Contact",i, contact);
      				
      				//reader.getCellData("TestData", "Contact", i);
      		//======================================================================================================
      				
      				
      		/*	//another city
      		        String selectLinkOpeninNewTab1 = Keys.chord(Keys.CONTROL,"t");
      		        driver.findElement(By.linkText("https://www.icicilombard.com/campaigns/health-insurance/health-insurance-mumbai")).sendKeys(selectLinkOpeninNewTab1);
      		        
      		        
      		        List<WebElement> rows1=driver.findElements(By.xpath("//table[@id='HospitalList']//tr"));
      		      	System.out.println("row size="+(rows1.size()-1));
      		      	int rowCount1=rows1.size();
      		      	
      		      	
      		      	
      		      //scroll the screen
      		      		JavascriptExecutor jse1 = (JavascriptExecutor)driver;
      		      		jse1.executeScript("window.scrollBy(0,950)");
      		      		//Thread.sleep(10000);
      		      		driver.findElement(By.id("Campaign1_C017_LnkSearch")).click();
      		      		//elementCSV.click();
      		      		
      		      		//For hospital
      		      		String beforexpath_hos1="//*[@id='HospitalList']/tbody/tr[";
      		      		String afterxpath_hos1="]/td[2]";
      		      		
      		      		
      		      		//Address
      		      		String beforexpath_Address1="//*[@id='HospitalList']/tbody/tr[";
      		      		String afterxpath_Address1="]/td[3]/span";


      		      		//City 
      		      		String beforexpath_City1="//*[@id='HospitalList']/tbody/tr[";
      		      		String afterxpath_City1="]/td[4]/span";
      		      		
      		      	
      		      		
      		      		//State
      		      		String beforexpath_state1="//*[@id='HospitalList']/tbody/tr[";
      		      		String afterxpath_state1="]/td[5]/span";
      		      		
      		      		//Contact
      		      		String beforexpath_contact1="//*[@id='HospitalList']/tbody/tr[";
      		      		String afterxpath_contact1="]/td[6]/span";
      		      		
      		      		
      		      		Xls_Reader reader1=new Xls_Reader("C:\\Users\\fidel\\eclipse-workspace\\LicHealth\\Driver\\Book1.xlsx");
      		          	reader1.addSheet("TestData2");
      		          	
      		          	reader1.addColumn("TestData2","Company Name");
      		          	reader1.addColumn("TestData2","Address");
      		          	reader1.addColumn("TestData2","City");
      		          	reader1.addColumn("TestData2","State");
      		          	reader1.addColumn("TestData2","Contact");
      		      			
      		      				for(int j=2;i<=rowCount1;i++)
      		      				{
      		      				//Hospital name
      		      				String actualxpath_hos1=beforexpath_hos1+i+afterxpath_hos1;

      		      				String hosame1= driver.findElement(By.xpath(actualxpath_hos1)).getText();
      		      				System.out.println(hosame1);
      		      				
      		      				reader1.setCellData("TestData2","Company Name",i, hosame1);
      		      				//reader.getCellData("TestData", "Company Name", i);
      		      				
      		      				//Address
      		      				String actualxpath_add1=beforexpath_Address1+i+afterxpath_Address1;
      		      				String Address1= driver.findElement(By.xpath(actualxpath_add1)).getText();
      		      				System.out.println(Address1);
      		      				reader1.setCellData("TestData2","Address",i, Address1);
      		      				//reader.getCellData("TestData", "Address", i);
      		      				
      		      				//City
      		      				String actualxpath_city1=beforexpath_City1+i+afterxpath_City1;
      		      				String City1= driver.findElement(By.xpath(actualxpath_city1)).getText();
      		      				System.out.println(City1);		
      		      				reader1.setCellData("TestData2","City",i, City1);
      		      				//reader.getCellData("TestData", "City", i);
      		      				
      		      				//State
      		      				String actualxpath_state1=beforexpath_state1+i+afterxpath_state1;
      		      				String state1= driver.findElement(By.xpath(actualxpath_state1)).getText();
      		      				System.out.println(state1);
      		      				reader1.setCellData("TestData2","State",i, state1);
      		      				
      		      				//reader.getCellData("TestData", "State", i);
      		      				//CONTACTNO
      		      				String actualxpath_cont1=beforexpath_contact1+i+afterxpath_contact1;
      		      				String contact1= driver.findElement(By.xpath(actualxpath_cont1)).getText();
      		      				System.out.println(contact1);
      		      				reader1.setCellData("TestData2","Contact",i, contact1);

      				
      				
      				}*/
      				
      				

    }
    }
}

