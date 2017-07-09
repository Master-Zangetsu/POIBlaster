package co.uk.nightshiftgroup.POIBlaster;

public class SheetLocation {
	
	private String sheetName;
	private int rowNumber;
	private int columnIdentifier;
	
	public SheetLocation(String sheetName, int rowNumber, int columnIdentifier) {
		this.sheetName = sheetName;
		this.rowNumber = rowNumber;
		this.columnIdentifier = columnIdentifier;
	}
	
	public String getSheetName() {
		return this.sheetName;
	}
	
	public int getRowNumber() {
		return this.rowNumber;
	}
	
	public int getColumnIdentifier() {
		return this.columnIdentifier;
	}
}
