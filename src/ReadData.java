

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFHyperlink;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class ReadData {

	public static void main(String[] args) throws IOException
	{
		File source = new File("C:\\Users\\Megha\\eclipse-workspace\\SeleniumPrograms\\TestData.xlsx");
        FileInputStream input = new FileInputStream(source);
        XSSFWorkbook wb = new  XSSFWorkbook(input);
        
       XSSFSheet sheet = wb.getSheetAt(0);
       
       int totalnumberrows = sheet.getLastRowNum();
       
       String text = sheet.getRow(0).getCell(0).getStringCellValue();
       System.out.println(text);
       
       double numericvalue = sheet.getRow(1).getCell(0).getNumericCellValue();
       System.out.println(numericvalue);
       
       XSSFCell datevalue = sheet.getRow(2).getCell(0);
       System.out.println(datevalue);
       
       XSSFHyperlink url  = sheet.getRow(3).getCell(0).getHyperlink();
       String url2 = url.getAddress();
       System.out.println(url2);
      
       System.out.println("...............");
       
       for(int i=0;i<=totalnumberrows;i++) 
       {
    	  XSSFCell result = sheet.getRow(i).getCell(0);
    	  System.out.println(result);
       }
     
	}

	
}
