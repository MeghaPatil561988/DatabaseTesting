

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.FilterOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFHyperlink;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WriteExcelData
{

	public static void main(String[] args) throws IOException
	{
		File source = new File("C:\\Users\\Megha\\Desktop\\TestData.xlsx");
        FileInputStream input = new FileInputStream(source);
        XSSFWorkbook wb = new  XSSFWorkbook(input);
        
       XSSFSheet sheet = wb.getSheetAt(0);
       
      sheet.getRow(0).getCell(1).setCellValue("pass");
      sheet.getRow(1).getCell(1).setCellValue("fail");
      sheet.getRow(2).getCell(1).setCellValue("125.55");
      sheet.getRow(3).getCell(1).setCellValue("05/08/2020");
      
      System.out.println("writing excel");
      
      FileOutputStream output = new FileOutputStream(source);
      wb.write(output);//u should write this line otherwise excel file get crashed
      wb.close(); 
     


	}

}
