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

public class Exc3 {
	public static void main(String[] args) throws IOException {
		File f = new File("C:\\Users\\skann\\eclipse-workspace\\MavenFifthExcel\\src\\test\\resources\\Book11.xlsx");
		FileInputStream fi = new FileInputStream(f);
		Workbook w = new XSSFWorkbook(fi);
		
		
		
		Sheet s = w.getSheet("Sheet1");
		
		for (int j = 0; j < s.getPhysicalNumberOfRows(); j++) {
			Row r = s.getRow(j);
		//	System.out.println(r);

			for (int i = 0; i < r.getPhysicalNumberOfCells(); i++) {
				Cell cell = r.getCell(i);
				System.out.println("before" + cell);

				int ct = cell.getCellType();
				System.out.println("after celltype=" + ct);

				
				
				
				
				if (ct == 1) {
					String val = cell.getStringCellValue();
					System.out.println(val);
					System.out.println("");
				} else {
					if (DateUtil.isCellDateFormatted(cell)) {
						Date a = cell.getDateCellValue();
						SimpleDateFormat b = new SimpleDateFormat("dd-MMM-yyyy");
						String val = b.format(a);
						System.out.println(val);
						System.out.println("");
					} else {
						double d = cell.getNumericCellValue();
						long e = (long) d;
						String val = String.valueOf(e);
						System.out.println(val);
						System.out.println("");
					}
				}

			}
		} // for

	}
}
