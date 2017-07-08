package co.uk.nightshiftgroup.POIBlaster;

import java.io.File;
import java.io.IOException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import CustomExceptions.IncorrectCellReferenceException;
import CustomExceptions.UnsupportedFileException;

public class POIBlasterOperation {

	private String operationType;
	private String fileName;
	private String workSheetName;
	private String cellRange;

	public static void main(String[] args) {

		String operationType = "";
		String fileName = "";
		String workSheetName = "";
		String cellRange = "";

		if (args.length > 0) {
			for (String arg : args) {
				if (arg.contains(":")) {
					try {

						String identifier = arg.substring(0, arg.indexOf(":"));
						String value = arg.substring(arg.indexOf(":") + 1, arg.length());

						switch (identifier) {
						case "OP":
							operationType = value;
							break;
						case "FILE":
							fileName = value;
							break;
						case "SHEETS":
							workSheetName = value;
							break;
						case "CELLS":
							cellRange = value;
							break;
						default:
							System.out.println("Unknown argument exception: " + identifier);
							System.out.println("Available arguments are: 'OP', 'FILE', 'SHEETS', 'CELLS'");
							break;
						}

					} catch (Exception e) {
						System.out.println("Malformed argument exception: '" + arg
								+ "'. For a full readout of acceptable params see git page or query with '/?'");
					}
				} else if (arg.equals("/?")) {
					System.out.println("Available arguments are: 'OP', 'FILE', 'SHEETS', 'CELLS'");
				} else {
					System.out.println("Malformed argument exception: '" + arg
							+ "'. For a full readout of acceptable params see git page or query with '/?'");
				}
			}

			try {
			System.out.println(new POIBlasterOperation(operationType, fileName, workSheetName, cellRange).execute());
			} catch (UnsupportedFileException e) {
				e.printStackTrace(System.out);
			} catch (InvalidFormatException e) {
				e.printStackTrace(System.out);
			} catch (IOException e) {
				e.printStackTrace(System.out);
			} catch (IncorrectCellReferenceException e) {
				e.printStackTrace(System.out);
			}

		}
	}

	public POIBlasterOperation(String operationType, String fileName, String workSheetName, String cellRange) {

		this.operationType = operationType;
		this.fileName = fileName;
		this.workSheetName = workSheetName;
		this.cellRange = cellRange;
	}

	public String execute() throws UnsupportedFileException, InvalidFormatException, IOException, IncorrectCellReferenceException {

		String interimReturnType = "Executed with params: " + "OP:" + operationType + " FILE:" + fileName + " SHEETS:"
				+ workSheetName + " CELLS:" + cellRange;

		if (operationType.equals("R") || operationType.equals(POIBlaster.READ)) {

			File targetFile = new File(fileName);
			if (targetFile.getName().contains(".xls")) {
				
				throw new UnsupportedFileException("Designation file is not currently supported.");

			} else if (targetFile.getName().contains(".xlsx")) {
				
				XSSFWorkbook workbook = new XSSFWorkbook(targetFile);
				
				String outputText = "";
				
				for (String sheetName : workSheetName.split(":")) {
					XSSFSheet sheet = workbook.getSheet(sheetName);
					
					outputText+= "<-----------------------{" + sheetName + "}----------------------->\n";
					
					if (cellRange.contains(":")) {
						
						String[] cellStrings = cellRange.split(":");
						
						CellReference startingCellRef = new CellReference(cellStrings[0]);
						CellReference endingCellRef = new CellReference(cellStrings[1]);
						
						int startingRow = startingCellRef.getRow();
						int endingRow = endingCellRef.getRow();
						
						int startingColumn = startingCellRef.getCol();
						int endingColumn = endingCellRef.getCol();
						
						for (int i = startingRow; i < endingRow; i++) {
							Row row = sheet.getRow(i);
							
							for (int x = startingColumn; x < endingColumn; x++) {
								Cell cell = row.getCell(x);
								
								String cellContents = cell.getStringCellValue();
								outputText+= cell.getAddress() + ": " + cellContents;
							}
						}			
					} else {
						if (cellRange.length() == 2) {
							
							CellReference ref = new CellReference(cellRange);
							
							Row row = sheet.getRow(ref.getRow());
							Cell cell = row.getCell(ref.getCol());
							
							String cellContents = cell.getStringCellValue();
							outputText+= cellRange + ": " + cellContents;
							
						} else {
							throw new IncorrectCellReferenceException("Incorrect cell reference " + cellRange);
						}
					}
				}

				workbook.close();
				
				return outputText;
			} else {
				throw new UnsupportedFileException("Designation file is not supported.");
			}
		}

		return interimReturnType;
	}
}
