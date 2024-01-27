package task_13;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WriteOperation {

	public static void main(String[] args) {

		try (XSSFWorkbook book = new XSSFWorkbook()) {
			XSSFSheet sheet = book.createSheet();

			Object[][] data = {

					{ "Name", "Age", "Email" }, { "John Doe", "30", "john@test.com" },
					{ "Jane", "28", "jane@test.com" }, { "Bob Smit", "35", "jackey@example.com" },
					{ "Swapnil", "37", "swapnil@example.com" }

			};

			int rowCount = 0;

			for (Object[] row1 : data) {

				XSSFRow row = sheet.createRow(rowCount++);

				int columnCount = 0;

				for (Object col : row1) {

					XSSFCell cell = row.createCell(columnCount++);

					if (col instanceof String) {

						cell.setCellValue((String) col);
					} else if (col instanceof Integer) {

						cell.setCellValue((Integer) col);

					}

					FileOutputStream output = null;
					try {
						output = new FileOutputStream("Question third.xlsx");
					} catch (FileNotFoundException e) {

						e.printStackTrace();
					}
					try {
						book.write(output);
					} catch (IOException e) {

						e.printStackTrace();
					}

				}

			}
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

	}

}