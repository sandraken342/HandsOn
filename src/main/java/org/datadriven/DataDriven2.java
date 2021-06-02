package org.datadriven;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Date;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class DataDriven2 {
	
	public static void main(String[] args) throws IOException {
		
		//File location
		File loc=new File("C:\\Users\\Prabu\\eclipse-workspace\\MavenProject\\Excel\\SeleniumBook.xlsx");
				
		//Reading values
		FileInputStream f=new FileInputStream(loc);
				
		//Specify workbook
		Workbook w=new XSSFWorkbook(f);
				
		//Specify sheet name
		Sheet s=w.getSheet("Sheet1");
		
		for(int i=0;i<s.getPhysicalNumberOfRows();i++)
		{
			Row r=s.getRow(i);
			for(int j=0;j<r.getPhysicalNumberOfCells();j++)
			{
				Cell c=r.getCell(j);
				System.out.println(c);
				
				int celltype=c.getCellType();
				System.out.println(celltype);
				
				if(celltype==1)
				{
					String str=c.getStringCellValue();
					System.out.println(str);
				}
				else if(celltype==0)
				{
					if(DateUtil.isCellDateFormatted(c))
					{
						Date getDate=c.getDateCellValue();
						SimpleDateFormat sim=new SimpleDateFormat("yyyy-mm-dd");
						String date=sim.format(getDate);
						System.out.println(date);
					}
					else
					{
						double getnum=c.getNumericCellValue();
						long l=(long)getnum;
						String number=String.valueOf(l);
						System.out.println(number);
					}
				}
			}
		}
				
				
	}

}
