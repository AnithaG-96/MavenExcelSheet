package org.excc;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Date;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

public class Exc2 {
public static void main(String[] args) throws IOException {
File f=new File("C:\\Users\\skann\\eclipse-workspace\\MavenFifthExcel\\src\\test\\resources\\sample-xls-file-for-testing.xls");
FileInputStream fi=new FileInputStream(f);
Workbook w=new HSSFWorkbook(fi);
Sheet s = w.getSheet("Sheet1");	
Row r = s.getRow(1);

for(int i=0;i<r.getPhysicalNumberOfCells();i++) {
Cell cell = r.getCell(i);
System.out.println(cell);

int ct = cell.getCellType();
System.out.println("celltype="+ct);
	
if(ct==1)	{
	String val = cell.getStringCellValue();
	System.out.println(val);
	System.out.println("");
}else {
	if(DateUtil.isCellDateFormatted(cell)) {
		Date dv = cell.getDateCellValue();
		SimpleDateFormat sf=new SimpleDateFormat("dd-MMM-yyyy");
		String val = sf.format(dv);
		System.out.println(val);
		System.out.println("");
	}else {
		double db = cell.getNumericCellValue();
		long ln=(long)db;
		String val = String.valueOf(ln);
		System.out.println(val);
		System.out.println("");
	}
	
}
	
	
	
	
	
}
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
}
}
