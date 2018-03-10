package generic;

import java.io.FileInputStream;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class Excel 
{
	public static String getValue(String path,String sheet,int r,int c) 
	{
		String v="";
		try
		{
			Workbook wb=WorkbookFactory.create(new FileInputStream(path));
			v=wb.getSheet(sheet).getRow(r).getCell(c).toString();
		}
		catch(Exception e)
		{
			
		}
		return v;
	}
	
	public static int getRowCount(String path,String sheet)
	{
		int rc=0;
		try
		{
			Workbook wb=WorkbookFactory.create(new FileInputStream(path));
			rc=wb.getSheet(sheet).getLastRowNum();
		}
		catch(Exception e)
		{
			
		}
		return rc;
	}
	public static int getColumnCount(String path, String sheet,int r)
	{
		int cc=0;
		try
		{
			Workbook wb=WorkbookFactory.create(new FileInputStream(path));
			cc=wb.getSheet(sheet).getRow(r).getLastCellNum();
		}
		catch(Exception e)
		{
			
		}
		return cc;
	}
	
	/*public static void main(String[] args) 
	{
		// TODO Auto-generated method stub
		String path="./data/input.xlsx";
		String sheet="Sheet1";
		String value =getValue(path,sheet,0,0);
		System.out.println(value);
		int rowCount=getRowCount(path,sheet);
		System.out.println(rowCount);
		int columnCount=getColumnCount(path,sheet,0);
		System.out.println(columnCount);
	}*/
}
