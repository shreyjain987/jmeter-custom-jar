package com.jmetercustom.functions;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.Calendar;
import java.util.Collection;
import java.util.LinkedList;
import java.util.List;

import org.apache.jmeter.engine.util.CompoundVariable;
import org.apache.jmeter.functions.AbstractFunction;
import org.apache.jmeter.functions.InvalidVariableException;
import org.apache.jmeter.samplers.SampleResult;
import org.apache.jmeter.samplers.Sampler;
import org.apache.jmeter.threads.JMeterVariables;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class GetCell extends AbstractFunction{
	public FileInputStream fis = null;
	public FileOutputStream fileOut = null;
	private XSSFWorkbook workbook = null;
	private XSSFSheet sheet = null;
	private XSSFRow row = null;
	private XSSFCell cell = null;

	private static final String MyFunctionName = "__GetCellValue";
	private static final List<String> desc = new LinkedList<String>();

	static {
		desc.add("Workbook Path");
		desc.add("Sheet Name");
		desc.add("Column Name");
		desc.add("Row Num");
	}

	private Object[] values;
	

	@Override
	public synchronized String execute(SampleResult previousResult, Sampler currentSampler)
			throws InvalidVariableException {
		String path = ((CompoundVariable) values[0]).execute().trim(); // parameter 1
		String sheetName = ((CompoundVariable) values[1]).execute().trim(); // parameter 2
		String colName = ((CompoundVariable) values[2]).execute().trim(); // parameter 3
		int rowNum = Integer.parseInt(((CompoundVariable) values[3]).execute());
		
		JMeterVariables vars = getVariables();

		try {
			fis = new FileInputStream(path);
			workbook = new XSSFWorkbook(fis);
			sheet = workbook.getSheetAt(0);
			fis.close();
		} catch (Exception e) {
			e.printStackTrace();
		}
		
		// new code for POI version - 4.1.2
		try {
			if (rowNum <= 0)
				return "";

			int index = workbook.getSheetIndex(sheetName);
			int col_Num = -1;
			if (index == -1)
				return "";

			sheet = workbook.getSheetAt(index);
			row = sheet.getRow(0);
			for (int i = 0; i < row.getLastCellNum(); i++) {
				// System.out.println(row.getCell(i).getStringCellValue().trim());
				if (row.getCell(i).getStringCellValue().trim().equals(colName.trim()))
					col_Num = i;
			}
			if (col_Num == -1)
				return "";

			sheet = workbook.getSheetAt(index);
			// row = sheet.getRow(rowNum - 1);
			row = sheet.getRow(rowNum);
			if (row == null)
				return "";
			cell = row.getCell(col_Num);

			if (cell == null)
				return "";

			// System.out.println(cell.getCellType().name());
			//
			if (cell.getCellType().name().equals("STRING")) {
				vars.put("CELL_TYPE_STRING", cell.getStringCellValue());
				return cell.getStringCellValue();
			}

			// if (cell.getCellType().STRING != null)

			// if(cell.getCellType()==Xls_Reader.CELL_TYPE_STRING)
			// return cell.getStringCellValue();
			else if ((cell.getCellType().name().equals("NUMERIC")) || (cell.getCellType().name().equals("FORMULA"))) {

				String cellText = String.valueOf(cell.getNumericCellValue());
				if (HSSFDateUtil.isCellDateFormatted(cell)) {
					// format in form of M/D/YY
					double d = cell.getNumericCellValue();

					Calendar cal = Calendar.getInstance();
					cal.setTime(HSSFDateUtil.getJavaDate(d));
					cellText = (String.valueOf(cal.get(Calendar.YEAR))).substring(2);
					cellText = cal.get(Calendar.DAY_OF_MONTH) + "/" + cal.get(Calendar.MONTH) + 1 + "/" + cellText;

					// System.out.println(cellText);

				}
				vars.put("CELL_TYPE_NUMERIC-FORMULA", cellText);
				return cellText;
			} else if (cell.getCellType().BLANK != null) {
				vars.put("CELL_TYPE_BLANK", "");
				return "";
			} else

				return String.valueOf(cell.getBooleanCellValue());

		} catch (Exception e) {

			e.printStackTrace();
			return "row " + rowNum + " or column " + colName + " does not exist in xls";
		}
	}
		

	@Override
	public String getReferenceKey() {
		return MyFunctionName;
	}

	@Override
	public void setParameters(Collection<CompoundVariable> parameters) throws InvalidVariableException {
		values = parameters.toArray();
	}

	@Override
	public List<String> getArgumentDesc() {
		return desc;
	}

}
