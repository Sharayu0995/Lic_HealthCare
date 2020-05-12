package com.practice.LicHealth;

/*import java.awt.AWTException;

import org.junit.Test;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;


public class AppTest 
{
 
	
	@Test
	public void typingInfo1(WebElement element) throws AWTException, InterruptedException
	{
	//Open Chrome driver
	System.setProperty("webdriver.chrome.driver","C:\\Users\\fidel\\eclipse-workspace\\LicHealth\\Driver\\chromedriver.exe");
	WebDriver driver= new ChromeDriver();
	driver.manage().window().maximize();
	driver.get("https://www.icicilombard.com/campaigns/health-insurance/health-insurance-pune");
	
		//scroll the screen
		JavascriptExecutor jse = (JavascriptExecutor)driver;
		jse.executeScript("window.scrollBy(0,250)");
		Thread.sleep(10000);
	}
//driver
	WebElement state=driver.findElement(By.id("sbSelector_82376309"));
	typingInfo1(state);

	WebElement city=driver.findElement(By.id("sbSelector_82376309"));
	typingInfo1(city);
	
	WebElement option=driver.findElement(By.id("sbSelector_28617778"));
	typingInfo1(option);
	
	try
	{
	
	if((element.getAttribute("id")=="state")&&(element.getAttribute("id")=="city")&&(element.getAttribute("id")=="option"))
	{
		driver.findElement(By.xpath("//a[@id='Campaign1_C017_LnkSearch']")).click();
	}
	
	}
	catch(Exception e)
	{
		System.out.println(e.getStackTrace());
	}
	

	}

}*/
