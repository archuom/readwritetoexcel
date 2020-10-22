package mypack;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Test4 {

	public static void main(String[] args) throws Exception
	{
		// Connect to existing excel file in project folder
		File f1 = new File("Book4.xlsx");
		
		//Take read permission on that file
		FileInputStream fi = new FileInputStream(f1);
		
		//access as a excel file
		Workbook rwb = WorkbookFactory.create(fi);
		
		Sheet sh1 = rwb.getSheet("sheet1");
		int nour = sh1.getPhysicalNumberOfRows();
	
		//Create new xlsx file
		File f2  = new File("ResultBook.xlsx");
		FileOutputStream fo = new FileOutputStream(f2);
		//HSSFWorkbook hwb = new HSSFWorkbook();// .xls workbook
		XSSFWorkbook wwb = new XSSFWorkbook();
		Sheet sh2 = wwb.createSheet("results");
		sh2.createRow(0).createCell(0).setCellValue("output");
		
		//get data from 2nd row(index=1) onwards in sheet1
		//1st row (index=0) has column names(input1, input2, output)
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
		//FileOutputStream fo = new FileOutputStream(f);
		
		wwb.write(fo);
		rwb.close();
		wwb.close();
		fi.close();
		fo.close();		
	}

}
