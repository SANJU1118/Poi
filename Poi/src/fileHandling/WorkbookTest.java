package fileHandling;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.testng.annotations.Test;

public class WorkbookTest
{
  @Test
  public void sheetTest1() throws EncryptedDocumentException, InvalidFormatException, IOException 
  {
	  FileInputStream fis = new FileInputStream("//home//sanjeev//Desktop//workbook.ods");  
	  Workbook wb = WorkbookFactory.create(fis);
	  Sheet sh = wb.getSheet("Sheet1");
	  int row = sh.getLastRowNum();
	  //String username = row.getCell(2).getStringCellValue();
	  System.out.println(row);
  }
}
