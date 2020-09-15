package com.poi.selenium;

import java.io.FileInputStream;
import java.util.concurrent.TimeUnit;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.testng.annotations.Test;
import org.testng.annotations.BeforeTest;

public class LoginWithPoi
{
	FileInputStream fis = null;
	Workbook wb = null;
	Sheet sh = null;
	XSSFCell cell = null;
	int rows = 0;
	
	public void getFile(String fileName, String sheetName) throws Exception
	{
		fis = new FileInputStream (fileName);
		wb = WorkbookFactory.create(fis);
		sh = wb.getSheet(sheetName);
		rows = sh.getPhysicalNumberOfRows();
	
	}
	
	@BeforeTest
	public void Test01() throws Exception 
	{
		getFile("Test.xls", "LoginXLS");
	}
	
	@Test
	public void Test02() throws Exception 
	{
		for(int i=1;i<rows;i++)
		{
		System.setProperty("webdriver.chrome.driver", "chromedriver.exe");
		WebDriver driver = new ChromeDriver();
		driver.get("file:///D:/Backup%20C%20drive/Desktop/Sagar/Course/Offline%20website/index.html");
		driver.manage().window().maximize();
		driver.manage().timeouts().pageLoadTimeout(20, TimeUnit.SECONDS);
		driver.manage().timeouts().implicitlyWait(20, TimeUnit.SECONDS);

		driver.findElement(By.id("email")).clear();
		driver.findElement(By.id("email")).sendKeys(sh.getRow(i).getCell(0).getStringCellValue());
	
		String pass = String.valueOf(sh.getRow(i).getCell(1).getNumericCellValue());
		driver.findElement(By.id("password")).clear();
		driver.findElement(By.id("password")).sendKeys(pass.substring(0, pass.indexOf('.')));
		
		driver.findElement(By.xpath("//button")).click();
		
		Thread.sleep(3000);
		driver.close();
		}
	}
}

