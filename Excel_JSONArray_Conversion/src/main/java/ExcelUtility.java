import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Iterator;
import java.util.Set;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.json.simple.JSONArray;
import org.json.simple.JSONObject;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Row.MissingCellPolicy;

public class ExcelUtility {

	public static int getColoumnIndex(String coloumnName,XSSFSheet sheet)
	{
		XSSFRow row=sheet.getRow(0);
		int count=0;
		Iterator<Cell> cellIterator=row.cellIterator();
		while(cellIterator.hasNext())
		{
			Cell cell=cellIterator.next();
			if(cell!=null)
			{
				if(cell.getStringCellValue().equalsIgnoreCase(coloumnName))
					{
					break;
					}
			}
			count++;				
		}
		
		for (Row row1 : sheet) {
			
		Cell cell=row1.getCell(0);
		if(cell!=null)
			System.out.println(cell.getStringCellValue());
		}
		
		return count;

	}
	
	
	public static void writeExcel(Set<JSONArray> set)
	{
		try {
			XSSFWorkbook workbook=new XSSFWorkbook(new FileInputStream("Login.xlsx"));
			XSSFSheet sheet=workbook.createSheet("JsonObjectOutput");
			
			int r=0;
			Iterator<JSONArray> iterator= set.iterator();
			while(iterator.hasNext())
			{
				String s=iterator.next().toString();
				sheet.createRow(r).createCell(0).setCellValue(s);		
				r++;
			}
			
			FileOutputStream fo=new FileOutputStream("Login.xlsx");
			workbook.write(fo);
			
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}
	
}
