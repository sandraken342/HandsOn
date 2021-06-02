package org.datadriven;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Insert {
	public static void main(String[] args) throws IOException {
		
		File loc=new File("C:\\Users\\Prabu\\eclipse-workspace\\MavenProject\\Excel\\Book1.xlsx");
				
		Workbook w=new XSSFWorkbook();
				
		Sheet s=w.createSheet("Java");
		
		Row r=s.createRow(0);
		
		Cell c=r.createCell(1);
		
		c.setCellValue("insert1");
		
		FileOutputStream fil=new FileOutputStream(loc);
		
		w.write(fil);
		
		System.out.println("Done");
	}

}
