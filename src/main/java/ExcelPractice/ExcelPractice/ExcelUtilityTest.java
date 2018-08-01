package ExcelPractice.ExcelPractice;

import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

public class ExcelUtilityTest {
	public static void main(String[] args) throws EncryptedDocumentException, InvalidFormatException, IOException {
		String[][] data = ExcelUtility.loadExcelData("MOCK_DATA.xlsx");
		for (int i = 0; i < data.length; i++) {
			for (int j = 0; j < data[i].length; j++) {
				System.out.print(data[i][j] + "\t");
			}
			System.out.println();
		}
		
		ExcelUtility.writeExcelData("mock2.xlsx", "data", data);
	}
}
