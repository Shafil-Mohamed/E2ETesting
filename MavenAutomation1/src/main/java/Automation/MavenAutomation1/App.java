package Automation.MavenAutomation1;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


/**
 * Hello world!
 */
public class App {
    public static void main(String[] args) throws IOException {
      FileInputStream fi=new FileInputStream("c:\\sample2.xlsx");
      XSSFWorkbook wb=new XSSFWorkbook(fi);            
      XSSFSheet ws=wb.getSheet("Sheet1"); 
      
      
      //XSSFWorkbook wbw = new XSSFWorkbook();
      //XSSFSheet sheet1 = wbw.createSheet("Result");
      //FileOutputStream fileOut = new FileOutputStream("D:\\Samplecopied.xlsx");
   
      
      
      int rc=ws.getPhysicalNumberOfRows() ;
      //System.out.println(rc);
      for(int i=0;i<rc;i++)
      {
    	  XSSFRow r=ws.getRow(i);
    	  //XSSFRow rw=sheet1.createRow(i);
    	
    	  int cc=r.getPhysicalNumberOfCells();
    	  for(int j=0;j<cc;j++)
    	  {
    		  XSSFCell c=r.getCell(j);
    		 // XSSFCell cw=rw.createCell(j);
    		  //cw.setCellValue(c.getStringCellValue());
    		  String k=c.getStringCellValue();
    		  if(k.equals("Error"))
    		  {
    		  }
    		  else
    		  {
    		  System.out.print(c.getStringCellValue()+"	"+(i+1)+" "+(j+1));
    		  }
    	  }
    	  System.out.println();
    	  
    	  }
    
      //wbw.write(fileOut);
      //fileOut.close();
      //wbw.close();
      wb.close();
    }
}

