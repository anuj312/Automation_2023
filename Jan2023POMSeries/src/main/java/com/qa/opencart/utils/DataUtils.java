package com.qa.opencart.utils;

import java.util.Hashtable;

public class DataUtils {
	

		public static Object[][] getTestData(String sheetName, String testName, ExcelReader xls) throws Exception

		{

		int testCaseNameRowNum = 1;

		while(!xls.getCellData(sheetName,0, testCaseNameRowNum).equalsIgnoreCase(testName))
		{
		if (xls.getCellData(sheetName, 0, testCaseNameRowNum).equalsIgnoreCase(""))
		throw new Exception();
		testCaseNameRowNum++;
		}
		System.out.println("Row number of the test case is -- " + testCaseNameRowNum);

		
		int totalColCount = 0;

		int colStartRowNum = testCaseNameRowNum + 1;

		while(! xls.getCellData(sheetName, totalColCount, colStartRowNum).equals(""))
		{

		totalColCount++;
		}

		System.out.println("Total columns for the test case are -- " +totalColCount);

               int dataStartRowNum = testCaseNameRowNum +2;
               int totalDataRows =0;

               while(! xls.getCellData(sheetName, 0, dataStartRowNum).equals("*"))
		       {

		       totalDataRows++;
                dataStartRowNum++;
		       }
		System.out.println("Total columns for the test case are -- " +totalDataRows);

                dataStartRowNum = testCaseNameRowNum +2;
                int finalRows = dataStartRowNum + totalDataRows;
		        Hashtable<String, String> table = null;
	        	Object[][] myData= new Object[totalDataRows][1];
		
		int i=0;
                for(int rNum = dataStartRowNum;rNum<finalRows;rNum++)
                 {

		           table = new Hashtable<String, String>();
		           
		           for(int cNum = 0; cNum < totalColCount; cNum ++)

		               {
	                	String data = xls.getCellData(sheetName, cNum, rNum);

		String key = xls.getCellData(sheetName, cNum, colStartRowNum) ;
		System.out.println(key+" ---- "+data);
		table.put(key, data);
		 }
		System.out.println(table);
		myData[i][0] = table; // Storing data of 1 row from table to object array
		System.out.println("------------------");
		i++;
      }
		

		return myData; 

		}
}

