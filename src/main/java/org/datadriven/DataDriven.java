package org.datadriven;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class DataDriven {
	
	public static void main(String[] args) throws IOException {
		//File location
		File loc=new File("C:\\Users\\Prabu\\eclipse-workspace\\MavenProject\\Excel\\SeleniumBook.xlsx");
		
		//Reading values
		FileInputStream f=new FileInputStream(loc);
		
		//Specify workbook
		Workbook w=new XSSFWorkbook(f);
		
		//Specify sheet name
		Sheet s=w.getSheet("Sheet1");
		
		//Scenario 1 -- Print a specific value
		Row r=s.getRow(1);
		
		Cell c=r.getCell(1);
		System.out.println(c);
		
		//Scenario 2 -- number of rows and columns 
		int rows=s.getPhysicalNumberOfRows();
		System.out.println("Available rows: "+rows);
		
		Row r2=s.getRow(0);
		int cells=r2.getPhysicalNumberOfCells();
		System.out.println("Available column from single row: "+cells);
		
		//Scenario 3 -- Print single row
		Row r3=s.getRow(0);
		for(int i=0;i<r3.getPhysicalNumberOfCells();i++)
		{
			Cell  cc=r3.getCell(i);
			System.out.println(cc);
		}
		
		//Scenario 4 -- Print entire sheet
		for(int i=0;i<s.getPhysicalNumberOfRows();i++)
		{
			Row r4=s.getRow(i);
			for(int j=0;j<r4.getPhysicalNumberOfCells();j++)
			{
				Cell c1=r4.getCell(j);
				System.out.println(c1);
				
				//CellType
				int celltype=c1.getCellType();
				System.out.println(celltype);
			}
		}		
	}

}
