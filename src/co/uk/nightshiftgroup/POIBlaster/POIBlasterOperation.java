package co.uk.nightshiftgroup.POIBlaster;

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
						System.out.println("Malformed argument exception: '" + arg + "'. For a full readout of acceptable params see git page or query with '/?'");
					}
				} else if (arg.equals("/?")) {
					System.out.println("Available arguments are: 'OP', 'FILE', 'SHEETS', 'CELLS'");
				} else {
					System.out.println("Malformed argument exception: '" + arg + "'. For a full readout of acceptable params see git page or query with '/?'");
				}
			}

			System.out.println(new POIBlasterOperation(operationType, fileName, workSheetName, cellRange).execute());
			
		}
	}

	public POIBlasterOperation(String operationType, String fileName, String workSheetName, String cellRange) {

		this.operationType = operationType;
		this.fileName = fileName;
		this.workSheetName = workSheetName;
		this.cellRange = cellRange;
	}
	
	public String execute() {
		
		String interimReturnType = "Executed with params: " + "OP:" + operationType + " FILE:" + fileName + " SHEETS:" + workSheetName + " CELLS:" + cellRange;
		
		if (operationType.equals("R") || operationType.equals(POIBlaster.READ)) {
			
		}
		
		return interimReturnType;
	}
}
