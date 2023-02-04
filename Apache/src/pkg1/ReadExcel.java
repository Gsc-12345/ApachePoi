package pkg1;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadExcel 
{
	public static void main(String[] args) throws IOException
	{
	 File f=new File("../Apache/Public Name.xlsx");   //Connection establish
	 FileInputStream fi=new FileInputStream(f);      //Stream Object
	 XSSFWorkbook xs=new XSSFWorkbook(fi);         //workbook object
	 XSSFSheet xt=xs.getSheetAt(0);               //Sheet object
	 
	 int r=xt.getPhysicalNumberOfRows();        //sheet number of rows fetch
	 for(int i=0;i<r;i++)                      //loop for rows
	 {
		 XSSFRow xr=xt.getRow(i);            //everytime it will create row obect
		 int c=xr.getPhysicalNumberOfCells();  // fetch the number of columns
		 for(int j=0;j<c;j++)                 //loops for columns
		 {
			 XSSFCell xc=xr.getCell(j);   //everytime it will create cell obj
			 System.out.println(xc.getStringCellValue());   //fetch the data of cell
		 }
	 }
		
	}

}
