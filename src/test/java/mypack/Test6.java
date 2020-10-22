//program to check the occurrence of a value say'B'  
//in the first row (say column name)
//|x |y |z | a | B| V | C| P| Q|......

package mypack;

import java.io.File;
import java.io.FileInputStream;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class Test6 {

	public static void main(String[] args) throws Exception
	{
		// Connect to existing excel file in project folder
		File f = new File("Book6.xlsx");
		//Take read permission on that file
		FileInputStream fi = new FileInputStream(f);
		//access as a excel file
		Workbook wb = WorkbookFactory.create(fi);
		
		Sheet sh = wb.getSheet("sheet1");
		int nour = sh.getRow(0).getLastCellNum();
		System.out.println("no of rows"+nour);
		
		//bubble sort searching technique
		for(int i=0;i< nour;i++)
		{
			//get the string value in each column of row 0
			String target = sh.getRow(0).getCell(i).getStringCellValue();
			if(target.equals("B"))
			{
				System.out.println ("column found at index" + (i+1));
				break;
			}
		}
		wb.close();
		fi.close();
		
	}

}
