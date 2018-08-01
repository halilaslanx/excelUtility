package ExcelPractice.ExcelPractice;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.WorkbookUtil;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelUtility {
	public static String[][] loadExcelData(String filename)
			throws EncryptedDocumentException, InvalidFormatException, IOException {
		return loadExcelData(filename, 0, "");
	}

	public static String[][] loadExcelData(String filename, String sheetName) throws EncryptedDocumentException, InvalidFormatException, IOException {
		return loadExcelData(filename, -1, sheetName);
	}
	
	public static String[][] loadExcelData(String filename, int sheetNumber) throws EncryptedDocumentException, InvalidFormatException, IOException {
		return loadExcelData(filename, sheetNumber, "");
	}
	
	private static String[][] loadExcelData(String filename, int sheetNumber, String sheetName)
			throws EncryptedDocumentException, InvalidFormatException, IOException {

		try (Workbook wb = WorkbookFactory.create(new FileInputStream(filename))) {
			Sheet sh = null;
			if (sheetNumber > -1)
				sh = wb.getSheetAt(sheetNumber);
			else 
				sh = wb.getSheet(sheetName);
			
			int numberOfRows = sh.getLastRowNum() + 1;
			int numberOfPhysRows = sh.getPhysicalNumberOfRows();
			String[][] sheetData = new String[numberOfPhysRows][];
			
			int i = 0;
			for (int xlIndex = 0; xlIndex < numberOfRows; xlIndex++) {
				// System.out.println(i);
				if (sh.getRow(xlIndex) != null) {
					String[] rowData = new String[sh.getRow(xlIndex).getLastCellNum()];
					for (int j = 0; j < rowData.length; j++) {
						// System.out.print(j + " ");
						if (sh.getRow(xlIndex).getCell(j) != null) 
							rowData[j] = sh.getRow(xlIndex).getCell(j).getStringCellValue();
						else
							rowData[j] = "";
					}
					// System.out.println();
					sheetData[i] = rowData;
					i++;
				}
			}
			return sheetData;
		}
	}
	
	public static void writeExcelData(String filename, String sheetName, String[][] data)
			throws EncryptedDocumentException, InvalidFormatException, IOException {

		try (FileOutputStream fos = new FileOutputStream(filename)) {
			Workbook wb = new XSSFWorkbook();
			Sheet sh = wb.createSheet(sheetName);
			
			for (int i = 0; i < data.length; i++) {
				Row thisRow = sh.createRow(i);
				for (int j = 0; j < data[i].length; j++) {
					thisRow.createCell(j).setCellValue(data[i][j]);
				}
			}
			wb.write(new FileOutputStream(filename));
			wb.close();
		}
	}
}
