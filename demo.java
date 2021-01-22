import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.FileWriter;
import java.io.IOException;
import java.util.Iterator;
import java.util.Map;
import java.util.Scanner;
import java.util.TreeMap;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.json.simple.JSONArray;
import org.json.simple.JSONObject;
import org.json.simple.parser.JSONParser;
import org.json.simple.parser.ParseException;


class write
{
	public void writemeth() 
	{
		JSONObject obj = new JSONObject();
		 obj.put("name","padma");
		  	obj.put("department","731116104033");
		  	obj.put("branch","tamil");
		  	obj.put("year",2020);

		  	try 
			{
				
		  		FileWriter file = new FileWriter("f:\\newfile.json");
				file.write(obj.toJSONString());
                 file.flush();
				JSONParser parser = new JSONParser();
			
						} 
		  	catch (IOException e) 
		  	{
							
							e.printStackTrace();
						}

	
	}
	


}
class read
{
	public void readmeth() throws org.json.simple.parser.ParseException
	{
	JSONParser parser = new JSONParser();
	
	try {
		Object obj = parser.parse(new FileReader("f:\\newfile.json"));
		

		JSONObject jsonObject = (JSONObject) obj;
		System.out.println(jsonObject);
		
		String name = (String) jsonObject.get("name");
		System.out.println(name);

		String department = (String) jsonObject.get("department");
		System.out.println(department);

		String branch = (String) jsonObject.get("branch");
		System.out.println(branch);

		long year = (long) jsonObject.get("year");
		System.out.println(year);
		

	} catch (FileNotFoundException e) 
	{
		e.printStackTrace();
	} catch (IOException e) {
		e.printStackTrace();
	}
	}
}
class excel
{
	public void methconvert()
	{
		
		XSSFWorkbook workbook = new XSSFWorkbook(); 
  
		Scanner sc=new Scanner(System.in);
		System.out.println("enter the excel sheet name");
		String s=sc.nextLine();
		XSSFSheet sheet = workbook.createSheet(s); 
				System.out.println("enter the column count");
				int n;
				n=sc.nextInt();
				System.out.println("enter the"+" "+s+"rows count");
				int n1;
				n1=sc.nextInt();
		Object[][] f =new Object[100][100];
		
		System.out.println("enter the colum name");
		
		String ss="";
		for(int j=0;j<=n;j++)
		{
			
			f[0][j]=sc.nextLine();
			ss+=f[0][j];
			
			//j++;
		}
		System.out.println("enter the "+ss+"for the "+s);
		for(int i=1;i<n1;i++)
		{
			System.out.println("enter the"+i+"st row");
			  for(int j=0;j<n1;j++)
		      {
				  f[i][j]=sc.nextLine();
		      }
			
		}
		int rownum = 0;
		
		for(Object[] player : f)
		{
		    Row row = sheet.createRow(rownum++);
		    
		    int colnum = 0;
		    for(Object value : player)
		    {
		        Cell  cell = row.createCell(colnum++);
		        if (value instanceof String) {
		            cell.setCellValue((String) value);
		        } else if (value instanceof Integer) {
		            cell.setCellValue((Integer) value);
		        }
		    }
		}
		try
		{
		   
		FileOutputStream fileOutputStream = new FileOutputStream(new File("d:\\convert.xlsx"));
         
		workbook.write(fileOutputStream);
		System.out.println("convert.xlsx written successfully on disk.");
		}
		catch (Exception e) 
		{
		    e.printStackTrace();
		} 
		
	}
}
public class demo {

	public static void main(String[] args) throws ParseException 
	{
		
		write w=new write();
		w.writemeth();
		read r=new read();
		r.readmeth();
		excel e=new excel();
		e.methconvert();
		
		
	
	}
}

