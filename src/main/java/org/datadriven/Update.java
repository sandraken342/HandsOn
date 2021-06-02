package org.datadriven;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Update {
	public static void main(String[] args) throws IOException {
		
		File loc=new File("C:\\Users\\Prabu\\eclipse-workspace\\MavenProject\\Excel\\Book1.xlsx");
		
		FileInputStream f=new FileInputStream(loc);
		
		Workbook w=new XSSFWorkbook(f);
		
		Sheet s=w.getSheet("Java");
		
		Row r=s.getRow(0);
		
		Cell c=r.getCell(1);
		
		String name=c.getStringCellValue();
		
		if(name.equals("insert1"))
		{
			c.setCellValue("update1");
		}
		
		FileOutputStream fil=new FileOutputStream(loc);
		w.write(fil);
		
		System.out.println("Done");
	}

}
