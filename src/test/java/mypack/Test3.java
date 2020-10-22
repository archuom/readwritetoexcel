package mypack;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class Test3 {

	public static void main(String[] args) throws Exception
	{
		// Connect to existing excel file in project folder
		
		File f = new File("Book3.xlsx");
		
		//Take read permission on that file
		FileInputStream fi = new FileInputStream(f);
		
		//access as a excel file
		Workbook wb = WorkbookFactory.create(fi);
		
		Sheet sh1 = wb.getSheet("sheet1");//existing
		//create new sheet with name 'results' and column heading as 'output'
		Sheet sh2 = wb.createSheet("results");
		sh2.createRow(0).createCell(0).setCellValue("output");
		int nour = sh1.getPhysicalNumberOfRows();
		
		//get data from 2nd row(index=1) onwards in sheet1
		//1st row (index=0) has column names(input1, input2)
		int x, y,z;
		for(int i=1;i<nour;i++)
		{
			x = (int)sh1.getRow(i).getCell(0).getNumericCellValue();
			y = (int)sh1.getRow(i).getCell(1).getNumericCellValue();
			z= x+y;
			//put the sum value in new sheet 
			sh2.createRow(i).createCell(0).setCellValue(z);
		}
		
		//Take write permission on that file
		FileOutputStream fo = new FileOutputStream(f);
		
		wb.write(fo);
		fi.close();
		fo.close();		
	}

}
