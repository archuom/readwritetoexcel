package mypack;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class Test5 {

	public static void main(String[] args) throws Exception
	{
		// Connect to existing excel file in project folder
		File f = new File("Book5.xlsx");
		//Take read permission on that file
		FileInputStream fi = new FileInputStream(f);
		//access as a excel file
		Workbook wb = WorkbookFactory.create(fi);
		
		Sheet sh = wb.getSheet("sheet1");
		int nour = sh.getPhysicalNumberOfRows();
		int nouc = sh.getRow(0).getLastCellNum();
		System.out.println("no of used rows"+nour);
		System.out.println("no of used columns"+nouc);
		//int rowsum,columnsum;
		//row sum
		for(int i=0;i<nour;i++)
		{
			int rowsum = 0;
			for(int j=0;j<nouc;j++)
			{
				int x = (int)sh.getRow(i).getCell(j).getNumericCellValue();
				System.out.println("x="+x);
				
				rowsum = rowsum+x;
				System.out.println("rowsum of "+  i+"="+rowsum );
			}
			sh.getRow(i).createCell(nouc).setCellValue(rowsum);			
		}
		
		//column sum
		for(int i=0;i<nouc;i++)//column wise
		{
			int columnsum = 0;
			for(int j=0;j<nour;j++)//row wise in each column
			{
				int x = (int)sh.getRow(j).getCell(i).getNumericCellValue();
				columnsum = columnsum+x;
			}
			if(i==0)
			{
				sh.createRow(nour).createCell(i).setCellValue(columnsum);
			}
			else
			{
				sh.getRow(nour).createCell(i).setCellValue(columnsum);
			}
		}
				
		//Take write permission on that file
		FileOutputStream fo = new FileOutputStream(f);
		
		wb.write(fo);
		wb.close();
		fi.close();
		fo.close();		
	}

}
