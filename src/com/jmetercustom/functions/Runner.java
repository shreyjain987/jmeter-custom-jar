package com.jmetercustom.functions;

public class Runner {

	public static void main(String[] args) {
		CellData reader = new CellData(
				"C:\\Users\\shrey_jain\\OneDrive - McGraw Hill Education\\Documents\\JMeter Projects\\SOA-JMeter\\gts-soa-jmeter\\CustomerInformation\\Data\\CustInfo_Data.xlsx");
		String sheetName = "Customer_Account_Details_Search";
		
		System.out.println(reader.getCellData(sheetName, "CustomerAccountId", 2));
//		String data = reader.getCellData(sheetName, 0, 2);
//		System.out.println(data);
		
//		int rowCount = reader.getRowCount(sheetName);
//		System.out.println("total rows: "+ rowCount);
//		
//		//reader.addColumn(sheetName, "status");
//		
//		if(! reader.isSheetExist("Regsitration")){
//			reader.addSheet("Regsitration");
//		}
//		reader.setCellData(sheetName, "status", 2, "PASS");
//		
//		System.out.println(reader.getColumnCount(sheetName));
		
		//reader.removeColumn("Regsitration", 0);
		
		System.out.println(reader.getCellData("Customer_Account_Details_Search", "CustomerAccountId", 3));
		System.out.println(reader.getCellData("Customer_Account_Details_Search", "OrderCountryCode", 2));

	}

}
