import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.util.Scanner;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Color;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class RC{
	
	public static void main(String args[]) throws Throwable
	{
		 RcComparetor();
		 try
	        { 
	           // Just one line and you are done !  
	           // /K : Carries out command specified by string 
	           Runtime.getRuntime().exec(new String[] {"cmd", "/Ksssss", "Start"}); 
	  
	        } 
	        catch (Exception e) 
	        { 
	            System.out.println("HEY Buddy ! U r Doing Something Wrong "); 
	            e.printStackTrace(); 
	        } 
		 
	}
	
	public static void RcComparetor() throws Throwable {
		
		
		File folder = new File("C:\\DRIVERS\\New folder");
        String[] files = folder.list();
        int count=1;
        for (String file : files) 
        {
        	test(file);
        	count++;
        	if(count==2)
        	break;
        }
				
	}
	public static void test (String file) throws Throwable
	{
		System.out.println(file);
		File fp=new File("C:\\DRIVERS\\New folder\\"+file);
    	FileInputStream fis = new FileInputStream(fp);
    	XSSFWorkbook wb = new XSSFWorkbook(fis);
    	System.out.println(wb.getNumberOfSheets());
    	for(int i=0;i<wb.getNumberOfSheets();i++)
    	{
    		getSheet(wb,i);
    		
    	}
    	
	}
	public static void getSheet(XSSFWorkbook wb,int i) throws FileNotFoundException
	{

		System.out.println(wb.getSheetName(i));
		XSSFSheet sheet=wb.getSheetAt(i);
		verifySheet(sheet);
	}
	public static void verifySheet(XSSFSheet sheet) throws FileNotFoundException
	{
		if(sheet.getSheetName().equals("IP"))
		{
			
			File folder = new File("C:\\DRIVERS\\New folder 1");
	        String[] files = folder.list();
	        for (String file : files) 
	        {	System.out.println("/////////////////////////////");
	        	Scanner sc2  = new Scanner(new File("C:\\DRIVERS\\New folder 1\\"+file));
	        	
	        	String color1="org.apache.poi.xssf.usermodel.XSSFColor@441ee8e";
	        	for(Row row : sheet) {
				
	        		for (Cell cell : row) {
	        			CellStyle cellStyle = cell.getCellStyle();
	        			Color color = cellStyle.getFillForegroundColorColor(); 
	        			if(color.toString().equals((color1)))
	        			{
	        				
	        				if(sheet.getRow(0).getCell(cell.getColumnIndex()).getStringCellValue().equals("OOPS"))	
	        				{
						
	        					while (sc2.hasNext()) {
	        						String s = sc2.next();
	        						if(s.substring(0,2).equals("AV"))
	        						{
	        							System.out.println("ONE");
	        							System.out.println(s.substring(0,2));
	        							System.out.println(s.substring(10,15).equals(cell.getStringCellValue()));
	        							System.out.println(s.substring(29,30));
	        							break;
	        						}
	        					}
						//System.out.println(cell.getStringCellValue());
	        				}
	        				else if(sheet.getRow(0).getCell(cell.getColumnIndex()).getStringCellValue().equals("OOPS3"))
	        				{
	        					while (sc2.hasNext()) {
	        					String s = sc2.next();
						
	        						if(s.substring(0,2).equals("AD"))
	        						{
	        							System.out.println(s.substring(0,2));
	        							System.out.println(s.substring(10,15).equals(cell.getStringCellValue()));
	        							System.out.println(s.substring(29,30));
	        							break;
	        						}
	        					}
				
	        				}
	        				else if(sheet.getRow(0).getCell(cell.getColumnIndex()).getStringCellValue().equals("OOPS4"))	
	        				{
						
	        					while (sc2.hasNext()) {
	        						String s = sc2.next();
	        						if(s.substring(0,2).equals("AC"))
	        						{							
	        							System.out.println(s.substring(0,2));
	        							System.out.println(s.substring(10,15).equals(cell.getStringCellValue()));
	        							System.out.println(s.substring(29,30));
	        							break;
	        						}
	        					}
						
	        				}
	        			}
	        		}
	        	}
	        }	
		}
	}
}

	

