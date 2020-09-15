package com.poi.selenium;

import java.io.FileInputStream;
import java.io.FileOutputStream;
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
import org.testng.Assert;
import org.testng.annotations.AfterTest;

public class PoiLoginWrite 
{
	FileInputStream fis=null;
	Workbook wb=null;
	Sheet sh=null;
	XSSFCell cell=null;
	int rows=0;
	
	public void getfile(String fileName, String sheetname) throws Exception
	{
		fis=new FileInputStream(fileName);
		wb =WorkbookFactory.create(fis);
		sh=wb.getSheet(sheetname);
		rows=sh.getPhysicalNumberOfRows();
	}
	
	@BeforeTest
	public void readTest() throws Exception
	{
		getfile("Test.xls","LoginXLS");
	}
	
	@Test
	public void test01() throws Exception{	
		for(int i=1;i<rows;i++) {
	System.setProperty("webdriver.chrome.driver", "chromedriver.exe");
	WebDriver driver= new ChromeDriver();
	driver.get("file:///D:/Backup%20C%20drive/Desktop/Sagar/Course/Offline%20website/index.html");
	driver.manage().window().maximize();
	driver.manage().timeouts().pageLoadTimeout(20, TimeUnit.SECONDS);//timeOut Exc
	driver.manage().timeouts().implicitlyWait(20, TimeUnit.SECONDS);//NoSuchEle
	driver.findElement(By.id("email")).clear();
	driver.findElement(By.id("email")).sendKeys(sh.getRow(i).getCell(0).getStringCellValue());
	String pass=String.valueOf((sh.getRow(i).getCell(1).getNumericCellValue()));//123456.0
	driver.findElement(By.id("password")).clear();
	driver.findElement(By.id("password")).sendKeys(pass.substring(0, pass.indexOf('.')));
	driver.findElement(By.xpath("//button")).click();
	if(driver.getTitle().equals("JavaByKiran | Dashboard")) 
		
		sh.getRow(i).createCell(2).setCellValue("PASS");
	else {
		sh.getRow(i).createCell(2).setCellValue("FAIL");
		
	}
	Thread.sleep(3000);
	driver.close();
		
	Assert.assertTrue(true);
	sh.getRow(i).createCell(3).setCellValue("PASS");
		
	}}

@AfterTest
public void writeExcel() throws Exception{
	FileOutputStream fos= new FileOutputStream("Test.xls");
	wb.write(fos);
	wb.close();
}
	
}
