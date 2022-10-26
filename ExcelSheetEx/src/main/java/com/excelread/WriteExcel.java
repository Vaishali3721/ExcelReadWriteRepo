package com.excelread;

import java.io.FileInputStream;
import java.io.FileOutputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class WriteExcel {
	
	public static void main(String[] args) throws Exception {
		Cell c=null;
		FileInputStream fis=new FileInputStream("TestData.xlsx");
		Workbook wb=WorkbookFactory.create(fis);
		
		if(wb.getSheet("Test")==null) {
			Sheet sh=wb.createSheet("Test");
			c=sh.createRow(2).createCell(3);
		}
		else {
			Sheet sh= wb.getSheet("Test");
			if(sh.getRow(2)==null) {
				c=sh.createRow(2).createCell(3);
			}else {
				c=sh.getRow(2).createCell(3);
			}
		}
		c.setCellValue("TheKiranAcademy");
		
		FileOutputStream fos=new FileOutputStream("TestData.xlsx");
		wb.write(fos);//to save the changes in the sheet
		wb.close();
		fos.close();
		System.out.println("Data added successfully");
	}

}
