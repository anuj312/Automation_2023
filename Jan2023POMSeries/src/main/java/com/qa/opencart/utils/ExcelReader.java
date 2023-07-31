package com.qa.opencart.utils;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Calendar;
import java.util.HashMap;
import java.util.Iterator;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelReader {
	
	

	private static final String TEST_DATA_SHEET_PATH = "./src/main/resources/testdata/OpenCartTestData.xlsx";
	
	int TotalSheets,k, TestScenario=8;
	

	String CurrentSheet;
	

	Row FirstRow;
	
		int totalSheets,i;
	
		public String path;
		public FileInputStream fis =null;
		public FileOutputStream fileout = null;
		private XSSFWorkbook workbook = null;
		private XSSFSheet sheet =null;
		private XSSFRow row= null;
		private XSSFCell cell= null;

		public ExcelReader(){

		}

		public ExcelReader(String sheetName){
		try
		{
		if(sheetName.equalsIgnoreCase("OpenCartTestData"))
		fis = new FileInputStream(System.getProperty("user.dir")+TEST_DATA_SHEET_PATH );
		else if(sheetName.equalsIgnoreCase("SmokeTestData"))
		fis = new FileInputStream(System.getProperty("user.dir")+TEST_DATA_SHEET_PATH );

		workbook = new XSSFWorkbook(fis);
		}
		catch(FileNotFoundException e){

		e.printStackTrace();
		}
		catch(IOException e){
         e.printStackTrace();
         }

		}

		public String getCellData(String sheetName,int colNum,int rowNum) {

		try {
		if (rowNum <= 0)
		return "";

		int index = workbook.getSheetIndex(sheetName);

		if (index == -1)
		return "";

		sheet = workbook.getSheetAt(index);
		row = sheet.getRow(rowNum - 1);
		if (row == null)

		return "";
		cell = row. getCell(colNum) ;
		if (cell == null)

		return "";

		if(cell.getCellType() == cell.getCellType().STRING)
		return cell.getStringCellValue();

		else if(cell.getCellType() == cell.getCellType().NUMERIC || cell.getCellType() == cell.getCellType().FORMULA){

		String cellText = String. valueOf(cell.getNumericCellValue());
		if (DateUtil.isCellDateFormatted(cell)) {

		// format in form of M/D/YY

		double d = cell.getNumericCellValue();

		Calendar cal = Calendar.getInstance();
		cal.setTime(DateUtil.getJavaDate(d));
		cellText =
		(String.valueOf(cal.get(Calendar.YEAR)));
		cellText = cal.get(Calendar.MONTH) + 1 + "/" +
		cal.get(Calendar.DAY_OF_MONTH) + "/" +
		cellText;
		}

		return cellText;

		}else if(cell.getCellType() == cell.getCellType().BLANK)
		return "";

		else
		return String.valueOf(cell.getBooleanCellValue());

		} catch (Exception e) {

		e.printStackTrace();

		return "row " + rowNum + " or column " + colNum + " does not exist in xls";
		}
		}

		public int getRowCount (String sheetName) {

		int index = workbook.getSheetIndex(sheetName) ;
		if(index==-1)

		return 0;
		else{

		sheet = workbook.getSheetAt(index) ;

		int number = sheet.getLastRowNum()+1;
		return number;

		 }

		}


		public boolean isSheetExist(String sheetName) {

		int index = workbook.getSheetIndex(sheetName) ;
		if(index==-1){

		index=workbook.getSheetIndex(sheetName.toUpperCase());
		if (index==-1)
		return false;
		else

		return true;
		}
		else
		return true;
		}




	public HashMap<Integer,HashMap<String,String>> GetTestData(String SheetName, String TestName) throws IOException {

	fis=new FileInputStream(TEST_DATA_SHEET_PATH );
	workbook = new XSSFWorkbook(fis);
    TotalSheets = workbook.getNumberOfSheets();
	HashMap<Integer, HashMap<String, String>> FullTestData = new HashMap<>();
	
	for(i=0;i<TotalSheets; i++)
	{
	CurrentSheet = workbook.getSheetName(i);

	if (CurrentSheet.equalsIgnoreCase(SheetName))
	{

	System.out.println("Go to Sheet -> "+SheetName);
	Sheet FeatureSheet = workbook.getSheetAt(i);
	Iterator<Row> rows = FeatureSheet.iterator();
	boolean testDataFlag=false;

	while(rows.hasNext())
	{
	Row CurrentRow = rows.next();
	if(CurrentRow.getCell(0).getStringCellValue().equalsIgnoreCase(TestName))
	{
	testDataFlag=true;
	System.out.println("Fetching test data for Scenario -> "+TestName) ;
	Row HeaderRow = rows.next();
	Row DataRow = rows.next();
	int TotalData=1;
	while(!DataRow.getCell(0).getStringCellValue().equalsIgnoreCase("*")){
	Iterator<Cell> HeaderCell = HeaderRow.cellIterator();
	Iterator<Cell> DataCell = DataRow.cellIterator();
	HashMap<String, String> DataRowTestData = new HashMap<>();
	while(HeaderCell.hasNext()) {
	Cell HeaderName = HeaderCell.next();
	if(!"".equals(HeaderName.getStringCellValue().trim())) {
	String sData = "";
	if(DataCell.hasNext()) {
	Cell Data = DataCell.next();
	
	switch(Data.getCellType())
	{
	
	case BOOLEAN:
	sData = String.valueOf(Data.getBooleanCellValue());
	break;
	
	case NUMERIC:
	if (DateUtil.isCellDateFormatted(Data)){
	//sData = sdf.format(Data.getDateCellValue()); 
		}
	else{
		
		sData = String.valueOf(Data.getNumericCellValue());
	
	try {
	if (Integer.parseInt(sData.split("\\.")[1]) > 0) {
	
		sData = String.valueOf(Data.getNumericCellValue());
	} else {

	sData = String.valueOf((long)Data.getNumericCellValue());

	}
	}catch(Exception e){
	sData = String. valueOf((long)Data.getNumericCellValue());
	}
	}
	break;
	
	case STRING:
    default:
	sData = Data.getStringCellValue();

	break;

	}
	}

	DataRowTestData.put(HeaderName.getStringCellValue(), sData);
	FullTestData.put(TotalData, DataRowTestData);

	}
	}
	DataRow=rows.next();
	TotalData++;
	}
	break;
	} 

	}
	if(!testDataFlag)System.out.println("No Data found for TestName ["+TestName+"] in TestData file");
	}
	}

	return FullTestData;
	}
	}

