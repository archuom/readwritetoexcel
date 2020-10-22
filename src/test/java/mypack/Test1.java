package mypack;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class Test1 {

	public static void main(String[] args) throws Exception
	{
		// Connect to existing excel file in project folder
		
		File f = new File("Book1.xlsx");
		
		//Take read permission on that file
		FileInputStream fi = new FileInputStream(f);
		
		//access as a excel file
		Workbook wb = WorkbookFactory.create(fi);
		
		Sheet sh = wb.getSheet("sheet1");
		
		int nour = sh.getPhysicalNumberOfRows();
		
		//get data from 2nd row(index=1) onwards in sheet1
		//1st row (index=0) has column names(input1, input2, output)
		int x, y,z;
		for(int i=1;i<nour;i++)
		{
			x = (int)sh.getRow(i).getCell(0).getNumericCellValue();
			y = (int)sh.getRow(i).getCell(1).getNumericCellValue();
			z= x+y;
			sh.getRow(i).createCell(2).setCellValue(z);
		}
		
		//Take write permission on that file
		FileOutputStream fo = new FileOutputStream(f);
		
		wb.write(fo);
		fi.close();
		fo.close();		
	}

}
