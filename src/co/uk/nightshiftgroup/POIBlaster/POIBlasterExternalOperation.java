package co.uk.nightshiftgroup.POIBlaster;

import org.apache.poi.ss.util.CellReference;
import CustomExceptions.MismatchedDelimitedValuesException;
import CustomExceptions.UnknownOperationException;

public class POIBlasterExternalOperation {

	public static void main(String[] args) {

		String operationType = "";
		String fileName = "";
		String workSheetNames = "";
		String cellRange = "";
		String cellValues = "";

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
							workSheetNames = value;
							break;
						case "CELLS":
							cellRange = value;
							break;
						case "VALUES":
							cellValues = value;
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
					System.out.println("Available arguments are: 'OP', 'FILE', 'SHEETS', 'CELLS', 'VALUES'");
				} else {
					System.out.println("Malformed argument exception: '" + arg
							+ "'. For a full readout of acceptable params see git page or query with '/?'");
				}
			}

			try {
				System.out.println(performExternalOperation(operationType, fileName, workSheetNames, cellRange, cellValues));
			} catch (Exception e) {
				e.printStackTrace(System.out);
			}
		}
	}

	private static String performExternalOperation(String operationType, String fileName, String workSheetName,
			String cellRange, String cellValues) throws Exception {

		String response = "";

		if (operationType.equals("R")) {

			String[] sheets = workSheetName.split(":");
			String[] sheetDelimitedCells = cellRange.split(":");

			if (sheets.length == sheetDelimitedCells.length) {
				for (int i = 0; i < sheets.length; i++) {

					String sheetName = sheets[i];
					String[] sheetSpecificCells = sheetDelimitedCells[i].split(",");

					response += "<-----" + sheetName + "----->\n";
					for (int x = 0; x < sheetSpecificCells.length; x++) {

						CellReference cellRef = new CellReference(sheetSpecificCells[x]);
						SheetLocation location = new SheetLocation(sheetName, cellRef.getRow(), cellRef.getCol());

						String readOperationResponse = new ReadOperation(fileName).setReadLocation(location)
								.executeReadOperation();
						response += sheetSpecificCells[x] + ":" + readOperationResponse + "\n";
					}
				}
			} else {
				throw new MismatchedDelimitedValuesException(
						"The number of sheet delimited cells does not match the number of sheets provided.");
			}
		} else if (operationType.equals("W")) {

			boolean writeOperationSuccesful = true;

			String[] sheets = workSheetName.split(":");
			String[] sheetDelimitedCells = cellRange.split(":");
			String[] sheetDelimitedValues = cellValues.split(":");

			if ((sheets.length == sheetDelimitedCells.length) && (sheets.length == sheetDelimitedValues.length)) {
				for (int i = 0; i < sheets.length; i++) {

					String sheetName = sheets[i];
					String[] sheetSpecificCells = sheetDelimitedCells[i].split(",");
					String[] sheetSpecificValues = sheetDelimitedValues[i].split(",");

					if (sheetSpecificCells.length == sheetSpecificValues.length) {
						for (int x = 0; x < sheetSpecificCells.length; x++) {

							String cell = sheetSpecificCells[x];
							String value = sheetSpecificValues[x];

							CellReference cellRef = new CellReference(cell);
							SheetLocation location = new SheetLocation(sheetName, cellRef.getRow(), cellRef.getCol());

							boolean writeOperationResponse = new WriteOperation(fileName).setWriteLocation(location)
									.setWriteValue(value).executeWriteOperation();

							if (!writeOperationResponse) {
								writeOperationSuccesful = false;
							}
						}
					} else {
						throw new MismatchedDelimitedValuesException(
								"The number of delimited cells and values for this sheet do not match.");
					}
				}
			} else {
				throw new MismatchedDelimitedValuesException(
						"The number of sheet delimited cells or values does not match the number of sheets provided.");
			}
			response += writeOperationSuccesful;
		} else {
			throw new UnknownOperationException("Unrecognised operation type: " + operationType);
		}

		return response;
	}
}
