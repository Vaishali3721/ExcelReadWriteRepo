package com.writeresultinexcel;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class WriteResultExcel {
	
	public static void WriteResult(int row, int col, String data) throws Exception {
		Cell c=null;
		Sheet sh=null;
		FileInputStream fis=new FileInputStream("loginDataProvider.xlsx");
		Workbook wb=WorkbookFactory.create(fis);
		
		if(wb.getSheet("login")==null) {
			 sh=wb.createSheet("login");
			c=sh.createRow(row).createCell(col);
		}
		else {
			sh=wb.getSheet("login");
			if(sh.getRow(row)==null) 
				c=sh.createRow(row).createCell(col);
			else
			c=sh.getRow(row).createCell(col);
		}
		c.setCellValue(data);
		FileOutputStream fos=new FileOutputStream("loginDataProvider.xlsx");
		wb.write(fos);
		fos.close();
	}

}
