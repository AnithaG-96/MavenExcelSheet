package org.excc;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Date;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebElement;
import org.test.base.BaseClass;

public class ExcelFifthDay extends BaseClass {
public static void main(String[] args) throws IOException {
	
	chromeLaunch();
	urlLaunch("https://www.facebook.com/");
	impWait(10);
	
	WebElement user = driver.findElement(By.id("email"));
	sendKeys(user,excel("Book1","Sheet1", 1, 0));
	
	WebElement pass = driver.findElement(By.id("pass"));
	sendKeys(pass,excel("Book1", "Sheet1", 1, 1));
	
	WebElement lg = driver.findElement(By.name("login"));
	click(lg);
	
	 
//		
//	File f=new File("C:\\Users\\skann\\eclipse-workspace\\MavenFifthExcel\\src\\test\\resources\\Book1.xlsx");
//	FileInputStream fi=new FileInputStream(f);
//	Workbook w=new XSSFWorkbook(fi);
//	Sheet s = w.getSheet("Sheet1");
//	Row row = s.getRow(1);
//	Cell cell = row.getCell(3);
//	
//	int celltype = cell.getCellType();                 //celltype type1,type0 num/date
//	System.out.println(celltype);
//	
//	if(celltype==1) {
//		String value = cell.getStringCellValue();
//		System.out.println(value);
//		
//	}else {
//		if(DateUtil.isCellDateFormatted(cell)) {     
//			Date dd = cell.getDateCellValue();
//			SimpleDateFormat sd=new SimpleDateFormat("dd-MMM-yyyy");
//			String value = sd.format(dd);
//			System.out.println(value);
//			
//		}else {
//			double db = cell.getNumericCellValue();
//			long ln=(long)db;
//			String value = String.valueOf(ln);
//		    System.out.println(value);
//		}
//	}
	
	
	
	
	
	
	
	
	
	
	
	
	
}
}
