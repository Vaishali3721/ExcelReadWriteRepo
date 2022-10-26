package com.writeresultinexcel;

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
import org.testng.Assert;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;

public class WriteResultInExcel {
	
	WebDriver driver=null;
	int count=1;
	@Test(dataProvider="loginData") 
	public void login(String uname, String pass) throws Exception {
		System.setProperty("webdriver.chrome.driver", "chromedriver.exe");
		driver=new ChromeDriver();
		driver.get("file:///C:/Users/donga/OneDrive/Desktop/Offline%20Website/index.html");
		driver.findElement(By.id("email")).sendKeys(uname);
		driver.findElement(By.id("password")).sendKeys(pass);
		driver.findElement(By.xpath("//button")).click();
		
		Assert.assertEquals(driver.getTitle(), "JavaByKiran | Dashboard");
		if(driver.getTitle().equals("JavaByKiran | Dashboard")) {
			WriteResultExcel.WriteResult(count, 2, "PASS");
			count++;
		}
		else {
			WriteResultExcel.WriteResult(count, 2, "FAIL");
			count++;
		}
	}
	
	@DataProvider
	public Object[][] loginData() throws Exception {
		
		DataFormatter df=new DataFormatter();
		FileInputStream fis=new FileInputStream("loginDataProvider.xlsx");
		Workbook wb=WorkbookFactory.create(fis);
		Sheet sh=wb.getSheet("login");
		int rows=sh.getLastRowNum();
		System.out.println(rows);
		String[][] data=new String[rows][2];
		
		for(int i=0; i<rows; i++) {
			
				Cell c1=sh.getRow(i+1).getCell(0);
				Cell c2=sh.getRow(i+1).getCell(1);
			data[i][0]=df.formatCellValue(c1);
			data[i][1]=df.formatCellValue(c2);
		}
		
		return data;
		
	}

}
