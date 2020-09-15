package com.jbk;

import java.io.FileInputStream;
import java.io.FileOutputStream;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.testng.annotations.Test;


public class CommonReadWrite 
{
	@Test
	public void test01() throws Exception
	{
		FileInputStream fis = new FileInputStream("CommonReadWrite.xls");
		Workbook wb = WorkbookFactory.create(fis);
	
		Sheet sh =  wb.getSheetAt(0);
		Cell cell = sh.getRow(0).getCell(0);
		System.out.println(cell.toString());
		
		sh.createRow(5).createCell(0).setCellValue("Writing in XLS file");
		sh.createRow(6).createCell(1).setCellValue("Enter value into status");
		
		FileOutputStream fos = new FileOutputStream("CommonReadWrite.xls");
		wb.write(fos);
		wb.close();
	}
}
