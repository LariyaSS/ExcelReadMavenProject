package excelRead;

import java.io.IOException;

import org.apache.poi.util.SystemOutLogger;

public class ExcelMain {

	public static void main(String[] args) throws IOException {
		String s1 = ExcelRead.readStringData(1, 0);
		System.out.println(s1);
		String s2 = ExcelRead.readIntegerData(1, 1);
		System.out.println(s2);

	}

}
