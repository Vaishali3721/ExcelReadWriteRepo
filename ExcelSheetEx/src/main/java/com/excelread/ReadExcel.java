package com.excelread;

import java.io.FileInputStream;
import java.io.FileNotFoundException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class ReadExcel {
	
	//This method is for to get the data of single row single column 
	public static String getCellData(int row, int Col) throws Exception {
		DataFormatter df=new DataFormatter();
		FileInputStream fis=new FileInputStream("TestData.xlsx");
		Workbook wb=WorkbookFactory.create(fis);
		Sheet sh=wb.getSheet("Login");
		Cell c=sh.getRow(row).getCell(Col);
		
		return df.formatCellValue(c);
		
	}
	
	public static void main(String[] args) throws Exception {
		
		DataFormatter df=new DataFormatter();
		
		FileInputStream fis=new FileInputStream("TestData.xlsx");
		Workbook wb=WorkbookFactory.create(fis);
		Sheet sh=wb.getSheet("Login");
		int rows=sh.getLastRowNum();
		
		for(int i=0; i<=rows; i++) {
			int cols = sh.getRow(i).getLastCellNum();
			for(int j=0;j<=cols; j++) {
				Cell c=sh.getRow(i).getCell(j);
				System.out.print(df.formatCellValue(c)+ "  ");
			}
			System.out.println();
		}
		
		//calling method to get the data of single row & col 
		System.out.println(ReadExcel.getCellData(0, 1));
		
	}

}
