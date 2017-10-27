package BaseFactory;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelFactory

{
public File f;
public FileInputStream fs;
public XSSFWorkbook wb;
	
	public ExcelFactory() throws IOException
	{
		f=new File("E:\\workspace\\userdata.xlsx");
	   fs=new FileInputStream(f);
		wb=new XSSFWorkbook(fs);
	}
	
	
	
	public int rows()
	{
	int a=	wb.getSheet("Sheet1").getPhysicalNumberOfRows();
	return a;
	}
	
	public String getcellvalue(String Sheet,int row,int col) throws IOException
	{
	String data=wb.getSheet(Sheet).getRow(row).getCell(col).getStringCellValue();
	return data;
	
	}
	
}
