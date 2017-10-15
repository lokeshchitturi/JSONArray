import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.HashSet;
import java.util.Iterator;
import java.util.LinkedHashMap;
import java.util.LinkedList;
import java.util.List;
import java.util.Map;
import java.util.TreeMap;
import java.util.Set;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.json.simple.JSONArray;
import org.json.simple.JSONObject;

public class Test {

	public static void main(String[] args) {
		
		try {
			XSSFWorkbook workbook=new XSSFWorkbook(new FileInputStream("Login.xlsx"));
			XSSFSheet sheet=workbook.getSheet("Sheet1");

			
			Integer count=0;
			int j=0;
			JSONArray array=new JSONArray();
			Set<JSONArray> set=new HashSet<JSONArray>();
			
			for(Row row:sheet)
			{
				// to skip header line
				if(row.getRowNum()==0)
					continue;
				
				Cell testStepCell=row.getCell(3);
				
				if(testStepCell!=null)
				{
										
				Map<String, Object>map=new TreeMap<String, Object>();
				map.put("index", count);
				map.put("step", count+1);
				map.put("data", row.getCell(3).getStringCellValue());
				map.put("result1", row.getCell(4).getStringCellValue());
				
				JSONObject obj=new JSONObject(map);
				count=count+1;
				array.add(obj);
				
				}	
				else
				{
					System.out.println(array);
					set.add(array);
					array=new JSONArray();
					count=0;
				}
			}			
			System.out.println(set);
			System.out.println(set.size());
			
			ExcelUtility.writeExcel(set);
			
			
			
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}
}
