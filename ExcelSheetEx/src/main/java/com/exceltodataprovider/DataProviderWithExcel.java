package com.exceltodataprovider;

import java.io.FileInputStream;
import java.io.FileNotFoundException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;

public class DataProviderWithExcel {
	
	WebDriver driver=null;
	
	@Test(dataProvider="loginData")
	public void test01(String uname, String pass) {
		
		System.setProperty("webdriver.chrome.driver", "chromedriver.exe");
		driver= new ChromeDriver();
		driver.get("file:///C:/Users/donga/OneDrive/Desktop/Offline%20Website/index.html");
		driver.findElement(By.id("email")).sendKeys(uname);
		driver.findElement(By.id("password")).sendKeys(pass);
		driver.findElement(By.xpath("//button")).click();
	}
	
	@DataProvider
	public Object[][] loginData() throws Exception{
		
		DataFormatter df=new DataFormatter();
		
		FileInputStream fis=new FileInputStream("loginDataProvider.xlsx");
		Workbook wb=WorkbookFactory.create(fis);
		Sheet sh= wb.getSheet("login");
		int rows=sh.getLastRowNum();
		String[][] data=new String[rows][2];
		
		for(int i=1; i<=rows;i++) {
			Cell c1=sh.getRow(i).getCell(0);
			Cell c2=sh.getRow(i).getCell(1);
			data[i-1][0]=df.formatCellValue(c1);
			data[i-1][1]=df.formatCellValue(c2);
			
		}
		return data;
		
		
	}

}
