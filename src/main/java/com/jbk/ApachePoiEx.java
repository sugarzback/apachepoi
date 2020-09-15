package com.jbk;

import java.io.FileInputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ApachePoiEx 
{
	private static XSSFWorkbook wb;

	@SuppressWarnings("deprecation")
	public static void main(String[] args) throws Exception 
	{
		FileInputStream fis = new FileInputStream("Test.xlsx");
		wb = new XSSFWorkbook(fis);
		XSSFSheet sh = wb.getSheet("Login");
		
		int row = sh.getPhysicalNumberOfRows();
		int col = sh.getRow(row-1).getPhysicalNumberOfCells();
		System.out.println(row+"  "+col);
		
		for (int i=0; i<row;i++)
		{
			for(int j=0;j<col; j++)
			{
				Cell cell = sh.getRow(i).getCell(j);
				//System.out.println(cell.getStringCellValue()+"   ");				
				
				if(cell.getCellType()==Cell.CELL_TYPE_STRING)
					System.out.println(cell.getStringCellValue()+"  ");
				
				if(cell.getCellType()== Cell.CELL_TYPE_NUMERIC)
					System.out.println(cell.getNumericCellValue()+"  ");
			}
			System.out.println();
		}
		
	}
}
