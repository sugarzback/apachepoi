package com.jbk;

import java.io.FileInputStream;
import java.io.FileOutputStream;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.annotations.Test;

public class ApachePoiWrite 
{
	@Test
	public void test01() throws Exception
	{
		FileInputStream fis = new FileInputStream("Test.xlsx");
		XSSFWorkbook wb = new XSSFWorkbook(fis);
		XSSFSheet sh = wb.getSheet("PoiWrite");
		sh.getRow(1).createCell(4).setCellValue("Testing");
		sh.createRow(3).createCell(0).setCellValue("Done");
		
		FileOutputStream fos = new FileOutputStream("Test.xlsx");
		wb.write(fos);
		wb.close();
	}
}
