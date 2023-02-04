package pkg1;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WriteExcel
{
	public static void main(String[] args) throws IOException
	{
		File f=new File("../Apache/Public Name1.xlsx");  //Connection establish
		FileOutputStream fo=new FileOutputStream(f);   //outputStream object
		XSSFWorkbook xs=new XSSFWorkbook();           //workbook object
		XSSFSheet xt=xs.createSheet("SheetA");      //sheet object
		for(int i=0;i<3;i++)                       //loops for rows
		{
			XSSFRow xr=xt.createRow(i);          //row object
			for(int j=0;j<3;j++)               //loop for cell(column)
			{
				XSSFCell xc=xr.createCell(j);    //cell object
				xc.setCellValue("Deepak");       //Setcell data
			}
		}
		xs.write(fo);   //will move the data from workbook to outputStream
		fo.flush();     //will move the data from outputStream to file
		fo.close();    //for saving it
		
	}

}
