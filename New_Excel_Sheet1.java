package task_13;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class New_Excel_Sheet1 {

	public static void main(String[] args) {

		try (XSSFWorkbook book = new XSSFWorkbook()) {
			@SuppressWarnings("unused")
			XSSFSheet sheet = book.createSheet();

			FileOutputStream output = null;
			try {
				output = new FileOutputStream("Sheet1.xlsx");
			} catch (FileNotFoundException e) {

				e.printStackTrace();
			}
			try {
				book.write(output);
			} catch (IOException e) {

				e.printStackTrace();
			}
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}

}
