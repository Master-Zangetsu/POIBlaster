package co.uk.nightshiftgroup.POIBlaster;

import java.io.File;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import CustomExceptions.UnsupportedFileException;

public class WriteOperation {

	private String filePath;
	private SheetLocation sheetLocation;
	private String writeValue;

	public WriteOperation(String filePath) {
		this.filePath = filePath;
	}

	public WriteOperation setWriteValue(String writeValue) {
		this.writeValue = writeValue;
		return this;
	}

	public WriteOperation setWriteLocation(SheetLocation sheetLocation) {
		this.sheetLocation = sheetLocation;
		return this;
	}

	public Boolean executeWriteOperation() throws Exception {

		File file = new File(this.filePath);

		if (!file.exists()) {
			file.getParentFile().mkdirs();
			file.createNewFile();
		}

		if (file.getName().contains(".xls")) {

			throw new UnsupportedFileException("Designated file is not currently supported.");

		} else if (file.getName().contains(".xlsx")) {

			XSSFWorkbook workbook = new XSSFWorkbook(file);
			XSSFSheet sheet = workbook.getSheet(sheetLocation.getSheetName());

			Row row = sheet.getRow(sheetLocation.getRowNumber());
			Cell cell = row.getCell(sheetLocation.getColumnIdentifier());

			cell.setCellValue(writeValue);
			workbook.close();

			return true;
		}

		throw new UnsupportedFileException("Designated file is not currently supported.");
	}
}
