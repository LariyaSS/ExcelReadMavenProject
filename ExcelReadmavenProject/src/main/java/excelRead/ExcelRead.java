package excelRead;

import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelRead {

	static FileInputStream f;
	static XSSFWorkbook w;
	static XSSFSheet sh;

	public static String readStringData(int row, int col) throws IOException // inorder to read string data
	{
		f = new FileInputStream("C:\\Users\\lariy\\OneDrive\\Desktop\\Student.xlsx"); // path of file
		w = new XSSFWorkbook(f); // to get workbook from f
		sh = w.getSheet("Sheet1");
		XSSFRow r = sh.getRow(row); // getRow(row)
		XSSFCell c = r.getCell(col);
		return c.getStringCellValue(); // return type is string
	}

	public static String readIntegerData(int row, int col) throws IOException // inorder to read int data
	{
		f = new FileInputStream("C:\\Users\\lariy\\OneDrive\\Desktop\\Student.xlsx"); // path of file
		w = new XSSFWorkbook(f); // to get workbook from f
		sh = w.getSheet("Sheet1");
		XSSFRow r = sh.getRow(row); // getRow(row)
		XSSFCell c = r.getCell(col);
		int value = (int) c.getNumericCellValue();// it return double cell value so inorder to convert from double to
													// int type casting is done
		return String.valueOf(value); // value is int but return type is string
	}

}
