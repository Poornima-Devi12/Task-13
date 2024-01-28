package task13;
import java.io.FileOutputStream;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Write {

	public static void main(String[] args) {
		 // Creating a new Excel workbook
		XSSFWorkbook book = new XSSFWorkbook();

		// Creating a new sheet in the above created workbook
		XSSFSheet sheet = book.createSheet("Sheet1");

		// Data to be written in the workbook
		String[][] data = { { "Name", "Age", "Email" }, 
				            { "John Doe", "30", "john@test.com" },
				            { "Jane Doe", "28", "john@test.com" },  
				            { "Bob Smith", "35", "jacky@example.com" },
				             { "Swapnil", "37", "swapnil@example.com" } };

		int rowCount = 0;

		// loop to iterate over row
		for (String[] row1 : data)
		{
			XSSFRow row = sheet.createRow(rowCount++);

			int columnCount = 0;

			// Inner loop to iterate over cell
			for (String col : row1) 
			{
				XSSFCell cell = row.createCell(columnCount++);
				cell.setCellValue(col);
			}

		}

		// Using try catch 
		try (FileOutputStream output = new FileOutputStream("Writexl.xlsx");) {
			book.write(output);

	// Closing the workbook
			book.close();

		System.out.println("Data were written successfully to the WritexlFile.xlsx: ");
		} 
		catch (Exception e) 
		{
			e.printStackTrace();
		}

	}

}
