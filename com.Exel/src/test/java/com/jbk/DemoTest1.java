package com.jbk;

import java.io.FileInputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.testng.annotations.Test;

public class DemoTest1 {
	Sheet sh=null;
	Row r=null;
	Cell c=null;
	
	
	@Test
	public void Test() throws Exception
	{
		
		FileInputStream fis = new FileInputStream("Test.xlsx");
		Workbook wb = WorkbookFactory.create(fis);
		Sheet sh = wb.getSheet("Sheet1");
		sh=wb.createSheet("Test");		
		r=sh.createRow(3);
		c=r.createCell(5);
	}
	//else
	{
		//Workbook wb;
		//sh =wb.getSheet("Test");
		if(sh.getRow(3)==null) 
			r=sh.createRow(3);
		    c=r.createCell(5);
	}
}
      //c.setCelValue("TheKiranAcademy");

