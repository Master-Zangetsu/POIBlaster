package co.uk.nightshiftgroup.POIBlaster;

import java.io.File;
import java.io.IOException;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import CustomExceptions.UnsupportedFileException;

public class ReadOperation {
	
	private String filePath;
	private SheetLocation sheetLocation;
	
	public ReadOperation(String filePath) {
		this.filePath = filePath;
	}
	
	public ReadOperation setReadLocation(SheetLocation sheetLocation) {
		this.sheetLocation = sheetLocation;
		return this;
	}
	
	public String executeReadOperation() throws Exception {
		
		File file = new File(this.filePath);
		
		if (file.getName().contains(".xls")) {

			throw new UnsupportedFileException("Designated file is not currently supported.");

		} else if (file.getName().contains(".xlsx")) {
			
			XSSFWorkbook workbook = new XSSFWorkbook(file);
			XSSFSheet sheet = workbook.getSheet(sheetLocation.getSheetName());
			
			Row row = sheet.getRow(sheetLocation.getRowNumber());
			Cell cell = row.getCell(sheetLocation.getColumnIdentifier());
			
			String value = cell.getStringCellValue();
			workbook.close();
			
			return value;
		}
		
		throw new UnsupportedFileException("Designated file is not currently supported.");
	}
}
