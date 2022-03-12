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

public class Exc5 {
public static void main(String[] args) throws IOException {
File f=new File("C:\\Users\\skann\\eclipse-workspace\\MavenFifthExcel\\src\\test\\resources\\excel-spreadsheet-examples-for-students.xlsx");
FileInputStream fi=new FileInputStream(f);
Workbook w=new XSSFWorkbook(fi);
Sheet s = w.getSheet("au-500");
for(int j=0;j<s.getPhysicalNumberOfRows();j++) {
Row r = s.getRow(j);
System.out.println(r);

for(int i=0;i<r.getPhysicalNumberOfCells();i++) {
Cell cell = r.getCell(i);	
System.out.println(cell);

int ctype = cell.getCellType();
System.out.println("celltype==="+ctype);	
	
if(ctype==1)	{
	String a = cell.getStringCellValue();
	System.out.println(a);
	System.out.println("");
}else {
if(DateUtil.isCellDateFormatted(cell))	{
	Date b = cell.getDateCellValue();
	SimpleDateFormat c=new SimpleDateFormat("dd-MMMM-yyyy");
	String d = c.format(b);
	System.out.println(d);
	System.out.println("");
} else {
	double e = cell.getNumericCellValue();
	long ln=(long)e;
	String g = String.valueOf(ln);
	System.out.println(g);
	System.out.println("");
}
	
}
	
}}//for	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
}
}
