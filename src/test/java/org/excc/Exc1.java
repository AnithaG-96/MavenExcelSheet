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

public class Exc1 {
public static void main(String[] args) throws IOException {
	
File f=new File("C:\\Users\\skann\\eclipse-workspace\\MavenFifthExcel\\src\\test\\resources\\Book1.xlsx");
FileInputStream fi=new FileInputStream(f);
Workbook w=new XSSFWorkbook(fi);
Sheet s = w.getSheet("Sheet1");
Row r = s.getRow(1);	
Cell cell = r.getCell(3);	
	
int ct = cell.getCellType();
System.out.println(ct);	
	
if(ct==1)	{
	String ss = cell.getStringCellValue();
	System.out.println(ss);
}else {
	if(DateUtil.isCellDateFormatted(cell)) {
		Date dc = cell.getDateCellValue();
		SimpleDateFormat ff=new SimpleDateFormat("dd-MMM-Y");
		String ss = ff.format(dc);
		System.out.println(ss);
	}else {
		double n = cell.getNumericCellValue();
		long ln=(long)n;
		String ss = String.valueOf(ln);
		System.out.println(ss);
	}
	
}
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
}
}
