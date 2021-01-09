package org.apache.poiooxml;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;

import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class DiagonalAddition 
{
	public static void main(String[] args) throws Exception
	{
		//Connect to Excel file
		File f=new File("diagonaladdition.xlsx");
		//Take read permission on that file
		FileInputStream fi=new FileInputStream(f);
		//Consider that file as excel file
		Workbook wb=WorkbookFactory.create(fi);
		//Go to Sheet
		Sheet sh=wb.getSheet("Sheet1");
		//Count of rows and columns
		int nour=sh.getPhysicalNumberOfRows();
		int nouc=sh.getRow(0).getLastCellNum();
		//Data driven starts from 2nd row(index=1)
		int sum=0;
		if((nour-1)==nouc)
		{
			for(int i=1;i<nour;i++)
			{
				for(int j=0;j<nouc;j++)
				{
					if(j==(i-1))
					{
						DataFormatter df=new DataFormatter();
						int value=Integer.parseInt(df.formatCellValue(sh.getRow(i).getCell(j)));
						sum=sum+value;
					}
				}
			}
			sh.createRow(nour).createCell(nouc).setCellValue(sum);
		}
		
		for(int i=0;i<=nouc;i++)
		{
			sh.autoSizeColumn(i);
		}
		
		
		//Save and close excel
		FileOutputStream fo=new FileOutputStream(f);
		wb.write(fo);
		fi.close();
		fo.close();
		wb.close();
	}
}
