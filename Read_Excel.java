package task_13;

import java.io.FileInputStream;
import java.io.IOException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Read_Excel {

	public static void main(String[] args) {
		String excelFilePath = "C://Users//q//eclipse-workspace//ExcelFileOperation//Question third.xlsx";

		try (FileInputStream fileInputStream = new FileInputStream(excelFilePath);
				Workbook workbook = new XSSFWorkbook(fileInputStream)) {

			// Assuming the data is present in the first sheet (index 0)
			Sheet sheet = workbook.getSheetAt(0);

			// Iterate through each row
			for (Row row : sheet) {
				// Iterate through each cell in the row
				for (Cell cell : row) {
					// Print the cell value to the console
					System.out.print(cell.toString() + "\t");
				}
				System.out.println(); // Move to the next line after each row
			}

		} catch (IOException e) {
			e.printStackTrace();
		}
	}
}
