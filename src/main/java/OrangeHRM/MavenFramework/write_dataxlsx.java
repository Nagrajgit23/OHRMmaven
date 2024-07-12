package OrangeHRM.MavenFramework;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
public class write_dataxlsx {
	public static void main(String[] args) throws IOException 
	{
		String path = "D:\\DemoFile.xlsx";
		FileInputStream fips = new FileInputStream(path);
		XSSFWorkbook wb = new XSSFWorkbook(fips);
		Sheet sheet1 = wb.getSheetAt(0);
		int lastRow = sheet1.getLastRowNum();
		for(int i=0; i<=lastRow; i++){
		Row row = sheet1.getRow(i);
		Cell cell = row.createCell(2);

		cell.setCellValue("Write into Excel column3");
		Cell cell4 = row.createCell(3);
		cell4.setCellValue("Write into Excel column4");

		}

		FileOutputStream fops = new FileOutputStream(path);
		wb.write(fops);
		fops.close();
		}

}
