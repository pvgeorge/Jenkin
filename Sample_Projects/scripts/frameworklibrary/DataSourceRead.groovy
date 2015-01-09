package scripts.frameworklibrary

import com.eviware.soapui.SoapUI
import com.eviware.soapui.model.support.TestRunListenerAdapter
import com.eviware.soapui.model.testsuite.TestRunContext
import com.eviware.soapui.model.testsuite.TestRunner
import jxl.*; 
import groovy.sql.*;
import java.util.StringTokenizer;
import org.apache.log4j.Logger 

/*
* 
* @author : Kalesh Rajappan
* Version 	Author			Date Modified		Comments
* 1.0		Kalesh Rajappan	08-Nov-2013			Added the methods setCurrentDataRow and setNextRow
* 1.1						22-Nov-2013			Added the method loadConfigurationProperties
* 1.2						31-Dec-2013			Added the method setLastDataRow
* 1.3		Kalesh Rajappan	12-Jan-2014			Added the methods readDataFromExcelSheet and testDataSheetLoop & 
												updated methods setNextDataRow,setCurrentDataRow, setLastDataRow
*/

class DataSourceRead {
/*
	* This methods reads the data data from the excel sheet and store it to the soapUi properties step
	* @param fileName	  		- path of the excel data sheet
	* @param sheetName  		- Sheet from which data needs to be read
	* @param rowNumber  		- current row number
	* @param propertyStepName  	- soapUi property step to which the value will be loaded
	* @param nextStepAtEndOfData- Next step to execute at end of data
	* @param exitAtEndOfData  	- boolean flag to indicate if the loop should continue / exit at end of data
	* @param testRunner  		- Instance of test runner
	* @param context  			- Instance of run context
	* @param log  				- Instance of log
	* Revision History
	* Author			Date Modified			Comments
	* Kalesh Rajappan	12-Jan-2014 			Initial Version
	*/
	def readDataFromExcelSheet(String fileName,String sheetName,Integer rowNumber,String propertyStepName,String nextStepAtEndOfData,boolean exitAtEndOfData, TestRunner testRunner,TestRunContext context,Logger log)   { 
		//load the excel workbook
		def Workbook workbook 		= Workbook.getWorkbook(new File(fileName));
		def Sheet sheet        	= workbook.getSheet(sheetName); 
		def colcount    		= sheet.getColumns();
		def rowCount            = sheet.getRows();
		def lastDataRow = rowCount;
		
		//	access the properties step
		def properties 			= testRunner.testCase.getTestStepByName(propertyStepName) ;  	
		def currentRow 			= rowNumber; //context.CurrentDataRow 
		
		//in case there is no rows, return error
		if (rowCount <= 1) 
		{
			log.error "readDataFromExcelSheet : No Test Data available in the data sheet " + sheetName;
			testRunner.gotoStepByName(nextStepAtEndOfData);
			return;
		}
		
		//Get the LastDataRow for the sheet.
		// LastDataRow is used to have more control over data looping. Value is set from the function setLastDataRow.
		// check if the value is null. 
		if (context.getProperty("DataSheet_" + sheetName + "_LastDataRow") == null)
		{
			context.setProperty("DataSheet_" + sheetName + "_LastDataRow",rowCount);
		}
		else
		{
			lastDataRow	= context.getProperty("DataSheet_" + sheetName + "_LastDataRow");
			if (lastDataRow > rowCount)
			{
				lastDataRow = rowCount;
			}
		}
		
		//in case the row number is passed as an argument, then use it as the current row
		if (rowNumber == 0)
		{
			//when row number is zero, take the row number from the context property
			//check if the context property is null. For the first time the context property will be null
			if(context.getProperty("DataSheet_" + sheetName + "_CurrentDataRow") == null)
			{
				//set the value for CurrentDataRow as 1
				context.setProperty("DataSheet_" + sheetName + "_CurrentDataRow",1);
					currentRow = 1;
			}
			else
			{
				currentRow 			= context.getProperty("DataSheet_" + sheetName + "_CurrentDataRow");
			}
		}
		else
		{
			//when the row number value is non zero use it.
			currentRow 			= rowNumber;
			context.setProperty("DataSheet_" + sheetName + "_CurrentDataRow",rowNumber);			
		}
		
		//if the currentRow is greater than the last data row
		if (currentRow >= lastDataRow)
		{
			//if it is master sheet, or exitAtEndOfData flag is set, exit the loop at the end of data.
			if (exitAtEndOfData == true)
			{
					log.warn "readDataFromExcelSheet: Current data row number " + currentRow + " is more than the number of data rows/ last data row in data sheet " + sheetName + " .Executing Finish Step";
					testRunner.gotoStepByName(nextStepAtEndOfData)	;
					return;
			}
			else
			{
					//if exitAtEndOfData is not true , reset the current row to 1
					// with this the looping will continue from first row
					currentRow = 1			
			}
		}			

		// loop through the columns. Add each item to the property ( 0 based row and column)
		log.info "Storing the values from row " + currentRow + " of data sheet " + sheetName + " to SoapUI properties step " + propertyStepName;
		for(int i = 0; i < colcount; i++) {
		
			Cell dataCell = sheet.getCell(i,currentRow);
			String  data = dataCell.getContents();
			//Column name will be used as the property name.
			Cell headerCell = sheet.getCell(i,0);
			String  header = headerCell.getContents();   		
			properties.setPropertyValue(header,data)  
			//h=h+1	
		}	
		
	}
	/*
	* This methods checks if all the rows are executed and controls the looping
	* @param nextStepForLooping - Next step to execute for looping
	* @param nextStepAtEndOfData- Next step to execute when the datasheet reaches end of data.
	* @param testRunner  		- Instance of test runner
	* @param context  			- Instance of run context
	* @param log  				- Instance of log
	* Revision History
	* Author			Date Modified			Comments
	* Kalesh Rajappan	12-Jan-2014				Initial Version
	*/
  	def testDataSheetLoop(String sheetName,String nextStepForLooping, String nextStepAtEndOfData,TestRunner testRunner,TestRunContext context,Logger log)	{
		//get the current row and Row count of the sheet
		def currentRow 	= context.getProperty("DataSheet_" + sheetName + "_CurrentDataRow");
		def lastDataRow 	= context.getProperty("DataSheet_" + sheetName + "_LastDataRow");
		
		if (currentRow >= lastDataRow - 1)
		{
			//if all the rows are executed, go to step finish
			log.info "testDataSheetLoop: Reached end of data.";
			testRunner.gotoStepByName(nextStepAtEndOfData);
			return;
		}
		else
		{
			//if all the rows are not executed, increment the row count and go to the step to read the next row
			context.TestCaseStatus = "PASS";
			context.TestStatusDesc =  "";
			context.setProperty("DataSheet_" + sheetName + "_CurrentDataRow",currentRow + 1);
			testRunner.gotoStepByName(nextStepForLooping);
			return;
		}
	
	}
/*
	* This methods is used to set the next data row for execution
	* @param testRunner - Instance of test runner
	* @param context  	- Instance of run context
	* @param log  		- Instance of log
	*
	* Revision History
	* Author			Date Modified			Comments
	* Kalesh Rajappan	08-Nov-2013				Initial Version
	*/
  	def setNextDataRow(String sheetName,String nextStepAtEndOfData,boolean exitAtEndOfData,TestRunner testRunner,TestRunContext context,Logger log)	{

		//get the current row and last data row 
		def lastDataRow            = context.getProperty("DataSheet_" + sheetName + "_LastDataRow");
		def currentRow 			= context.getProperty("DataSheet_" + sheetName + "_CurrentDataRow");
		
		//check if current row is the last row
		if (currentRow >= lastDataRow  - 1)
		{
			if (exitAtEndOfData == true)
			{
				log.warn "setNextDataRow : Currently executing the last row. Executing step : " + nextStepAtEndOfData;
				testRunner.gotoStepByName(nextStepAtEndOfData);
				return;
			}
			else
			{
				//Reset the value to 1
				context.setProperty("DataSheet_" + sheetName + "_CurrentDataRow", 1);
			}
		}
		else
		{
			//increment the current row to 1 to make next row active
			context.setProperty("DataSheet_" + sheetName + "_CurrentDataRow",currentRow + 1);
			return;
		}
	
	}
	/*
	* This methods sets the active data row
	* @param testRunner - Instance of test runner
	* @param context  	- Instance of run context
	* @param log  		- Instance of log
	*
	* Revision History
	* Author			Date Modified			Comments
	* Kalesh Rajappan	08-Nov-2013				Initial Version
	*/
	def setCurrentDataRow(Integer rowNumber,String sheetName,String nextStepAtEndOfData,TestRunner testRunner,TestRunContext context,Logger log)   { 
	
		//get the last data row
		def lastDataRow         = context.getProperty("DataSheet_" + sheetName + "_LastDataRow");
		//def currentRow 			= context.getProperty("DataSheet_" + sheetName + "_CurrentDataRow")
		
		if (lastDataRow == null)
		{
			// set the active row
			context.setProperty("DataSheet_" + sheetName + "_CurrentDataRow",rowNumber);
			return;
		}
		//if the last row has a value, check if the input row number is greater than row count
		if (rowNumber >= lastDataRow)
		{
			log.error "setCurrentDataRow : Input data row number " + rowNumber + " is more than the number of last data row in datasheet " + sheetName+".";//Executing Finish Step"
			testRunner.gotoStepByName(nextStepAtEndOfData);
			return;
		}
		else
		{
			// set the active row
			context.setProperty("DataSheet_" + sheetName + "_CurrentDataRow",rowNumber);			
		}

	}
	/*
	* This methods sets the last data row
	* @param rowNumber 			- The rownumber to be set as the last row
	* @param sheetName 			- Name of the sheet
	* @param nextStepAtEndOfData- Name of the step to be executed if end of data has reached
	* @param testRunner 		- Instance of test runner
	* @param context  			- Instance of run context
	* @param log  				- Instance of log
	* Revision History
	* Author			Date Modified			Comments
	* Jessy Jose				31-Dec-2013				Initial Version
	*											This method will allow to set the last row only for a MasterDataSheet
	*/
	def setLastDataRow(Integer rowNumber,String sheetName,TestRunner testRunner,TestRunContext context,Logger log)   { 
	
		//get the current row and rowcount
		//def currentRow 			= context.getProperty("DataSheet_" + sheetName + "_CurrentDataRow");
		//def lastDataRow         = context.getProperty("DataSheet_" + sheetName + "_LastDataRow");
		
		context.setProperty("DataSheet_" + sheetName + "_LastDataRow",rowNumber + 1);
		log.info "setLastDataRow : setting the last data row for sheet " + sheetName + " as " + rowNumber
		
		/* With the changes 
		//check if the input row number is greater than cow count
		if (rowNumber >= lastDataRow)
		{
			log.error "Input data row number " + rowNumber + " is more than the number of data rows in datasheet "+sheetName+"."//Executing Finish Step"
			testRunner.gotoStepByName(nextStepAtEndOfData)
		}
		else
		{
			if (sheetName == context.MasterDataSheetName)
			{
				context.setProperty("DataSheet_" + sheetName + "_LastDataRow",rowNumber+1)			
			}
		}
		*/
	}
	
	
	/* ********************************************************************************************************
	* Function Name 	: loadExcelDataSheet
	* Description		: Loads the excel sheet from the specified location and store it in the TestRunContext
	*	This function is no longer required. Created a new function readDataFromExcelSheet to avoid storing the Datasheet into TestRunContext. 
	* @param filePath  		- location of the excel file
	* @param masterDataSheet- Name of the master sheet.
	* @param testRunner  	- Instance of test runner
	* @param context  		- Instance of run context
	* @param log  			- Instance of log
	*
	* Revision History
	* Author			Date Modified			Comments
	* Kalesh Rajappan	19-Sep-2013				Initial Version
	* Kalesh Rajappan	09-Dec-2013				Added input parameter mastersheet. Added logic to load multiple sheets
	* Kalesh Rajappan	31-Dec-2013				Added logic to set the LastDataRow of masterDataSheet
	************************************************************************************************************/
	
	def loadExcelDataSheet(String filePath, String masterDataSheet, TestRunner testRunner,TestRunContext context,Logger log)	{
		//Load the file
		Workbook workbook 		= Workbook.getWorkbook(new File(filePath))
		def Sheet[] sheets 		= workbook.getSheets()
		def sheetCount          = sheets.length;
		def masterSheetExist 	= false
		
		//Store the project path in a soapUi project property. This property can be used later in other steps
		def groovyUtils = new com.eviware.soapui.support.GroovyUtils(context)
		testRunner.testCase.testSuite.project.setPropertyValue( "projectPath",groovyUtils.projectPath)
		
		//Check if there are sheets in the workbook
		if (sheetCount < 1 )
		{
			log.error "No Test Data available in the workbook " + filePath
			testRunner.gotoStepByName("Finish")
			
		}
		else
		{	
			//Store the workbook to context variable 
			context.WorkBook = workbook
			//set the current row as 1
			context.CurrentDataRow = 1
			//Setting the Status as PASS by default. In case there is any error in the subsequent steps status will be updated
			context.TestCaseStatus = "PASS"
			context.TestStatusDesc =  ""

			log.info "Loaded test data from file " + filePath + ". Total number of worksheets : " + sheets.length

			for(int i = 0; i < sheetCount; i++) {
				//get each work sheet
				def Sheet sheet          = workbook.getSheet(i);
				//get the row count and name of the sheet
				def rowCount = sheet.getRows();
				def sheetName = sheet.getName()
				//load the Row count of the sheet to a context variable
				//context.setProperty("DataSheet_" + sheetName + "_RowCount",rowCount)
				//Initialize the current raw for the sheet (<ToDo> what is there is no data on the sheet)
				context.setProperty("DataSheet_" + sheetName + "_CurrentDataRow",1)
				log.info "DataSheet "+sheetName+" RowCount is " + rowCount
				
				if ( sheetName == masterDataSheet)
				{
					masterSheetExist 				= true
					context.MasterDataSheetRowCount = rowCount
					context.MasterDataSheetName 	= masterDataSheet
					//Set the rowcount to LastDataRow
					context.setProperty("DataSheet_" + sheetName + "_LastDataRow",rowCount)
					log.info "Master data sheet loaded. Row count : " + rowCount
				}
			}
			if (masterSheetExist == false)
			{
					log.error "Could not find the master data sheet " + masterDataSheet + " in the file " + filePath
					testRunner.gotoStepByName("Finish")
			}
		}

	}

	/*
	* This methods transfers the data from current row of the sheet to the soapUI properties
	*	This function is no longer required. Created a new function readDataFromExcelSheet to avoid storing the Datasheet into TestRunContext. 
	* @param rowNumber  		- current row number
	* @param sheetName  		- Sheet from which data needs to be read
	* @param propertyStepName  	- soapUi property step to which the value will be loaded
	* @param nextStepAtEndOfData- Next step to execute at end of data
	* @param exitAtEndOfData  	- boolean flag to indicate if the loop should continue / exit at end of data
	* @param testRunner  		- Instance of test runner
	* @param context  			- Instance of run context
	* @param log  				- Instance of log
	* Revision History
	* Author			Date Modified			Comments
	* Kalesh Rajappan	12-Sep-2013				Initial Version
	* Kalesh Rajappan	31-Dec-2013				Added logic to get the LastDataRow of masterDataSheet
	*/
	def readFromDataSheet(Integer rowNumber,String sheetName,String propertyStepName,String nextStepAtEndOfData,boolean exitAtEndOfData, TestRunner testRunner,TestRunContext context,Logger log)   { 

		def workbook 			= context.WorkBook
		def Sheet sheet        	= workbook.getSheet(sheetName); 
		//def Sheet sheet 		= context.TestDataSheet
		def colcount    		= sheet.getColumns();
		def rowCount            = sheet.getRows();
		//In case the sheet is masterDataSheet, assign the lastRow value to rowcount
		if (sheetName == context.MasterDataSheetName)
		{
			rowCount			= context.getProperty("DataSheet_" + sheetName + "_LastDataRow")
		}
				
		def properties 			= testRunner.testCase.getTestStepByName(propertyStepName)   	
		def currentRow 			= rowNumber //context.CurrentDataRow 
		
		if (rowCount <= 1) 
		{
			log.error "No Test Data available in the datasheet " + sheetName;
			testRunner.gotoStepByName(nextStepAtEndOfData);
			return;
		}
		
		//in case the row number is passed as an argument, then use it as the current row
		if (rowNumber == 0)
		{
			//when row number is zero, take the row number from the context property
			currentRow 			= context.getProperty("DataSheet_" + sheetName + "_CurrentDataRow")//context.CurrentDataRow
		}
		else
		{
			//when the row number value is non zero use it.
			currentRow 			= rowNumber 
			context.setProperty("DataSheet_" + sheetName + "_CurrentDataRow",rowNumber)
			//if the sheet is the master sheet, counter will be set as the value in the input
			//this will override the the current looping
			if (sheetName == context.MasterDataSheetName)
			{
				context.CurrentMasterDataSheetRow = rowNumber			
			}
		}
		
		//if the currentRow is greater than the total number of rows of the sheet
		if (currentRow >= rowCount)
		{
			//if it is master sheet, or exitAtEndOfData flag is set, exit the loop at the end of data.
			if ((sheetName == context.MasterDataSheetName) || (exitAtEndOfData == true))
			{
					log.warn "Current data row number " + currentRow + " is more than the number of data rows in data sheet " + sheetName + " .Executing Finish Step";
					testRunner.gotoStepByName(nextStepAtEndOfData)	;
					return;
			}
			else
			{
					//if the current sheet is not the master sheet or exitAtEndOfData is not true , reset the current row to 1
					currentRow = 1			
			}
		}			

		// loop through the columns. Add each item to the property ( 0 based row and column)
		log.info "Storing the values from row " + currentRow + " of data sheet " + sheetName + " to SoapUI properties step " + propertyStepName;
		for(int i = 0; i < colcount; i++) {
		
			Cell dataCell = sheet.getCell(i,currentRow);
			String  data = dataCell.getContents();
			//Coulumn name will be used as the property name.
			Cell headerCell = sheet.getCell(i,0);
			String  header = headerCell.getContents();   		
			properties.setPropertyValue(header,data)  
			//h=h+1	
		}	
		//increment the current row of the datasheet by 1
		//context.setProperty("DataSheet_" + sheetName + "_CurrentDataRow",currentRow + 1)
		//if master data sheet , increment the counter for the MasterDatasheet by 1
		//if (sheetName == context.MasterDataSheetName)
		//{
		//	context.CurrentMasterDataSheetRow = currentRow + 1		
		//}

	}

	/*
	* This methods checks if all the rows are executed and controls the looping
	*	This function is no longer required. Created a new function testDataSheetLoop. Use this function along with the function readDataFromExcelSheet
	* @param nextStepForLooping - Next step to execute for looping
	* @param nextStepAtEndOfData- Next step to execute when the datasheet reaches end of data.
	* @param testRunner  		- Instance of test runner
	* @param context  			- Instance of run context
	* @param log  				- Instance of log
	* Revision History
	* Author			Date Modified			Comments
	* Kalesh Rajappan	12-Sep-2013				Initial Version
	* Kalesh Rajappan	31-Dec-2013				Added logic to get the LastDataRow of masterDataSheet
	*/
  	def testDataLoop(String nextStepForLooping, String nextStepAtEndOfData,TestRunner testRunner,TestRunContext context,Logger log)	{
		//get the Row count and current row of master datasheet.
		//Test case looping is controlled based on the data in the MasterDataSheet.
		def currentRow 	= context.CurrentMasterDataSheetRow
		//def rowCount 	= context.MasterDataSheetRowCount
		def rowCount 	= context.getProperty("DataSheet_" + context.MasterDataSheetName + "_LastDataRow")

		if (currentRow >= rowCount - 1)
		{
			//if all the rows are executed, go to step finish
			testRunner.gotoStepByName(nextStepAtEndOfData);
			return;
		}
		else
		{
			//if all the rows are not executed, increment the row count and go to the step to read the next row
			context.TestCaseStatus = "PASS"
			context.TestStatusDesc =  ""
			context.CurrentMasterDataSheetRow = currentRow + 1;
			context.setProperty("DataSheet_" + context.MasterDataSheetName + "_CurrentDataRow",currentRow + 1)
			testRunner.gotoStepByName(nextStepForLooping);
			return;
		}
	
	}
	
	/* ********************************************************************************
	* Function Name 	: loadConfigurationProperties
	* Description		: Loads the property values to the project property
	* @param filePath  	- path for the config file
	* @param testRunner	- Instance of test runner
	* @param log  		- Instance of log
	*
	* Revision History
	* Author			Date Modified			Comments
	* Kalesh Rajappan	22-Nov-2013				Initial Version
	********************************************************************************* */
	def loadConfigurationProperties(String filePath, TestRunner testRunner,Logger log)	{
		
		Properties prop = new Properties()  
		//Load the properties file            
		prop.load(new FileInputStream(filePath));

		//add the values to the project property
		//get each of the property and set it as a property in the run context
		prop.each { key, value ->
			testRunner.testCase.testSuite.project.setPropertyValue("${key}","${value.value}")
		}
		log.info "Project properties set"
	}
	
	/*
	* This methods checks if all the rows are executed and controls the looping
	* @param filePath  			- File to read
	* @param propertyStepName  	- Name of the soapUi properties step
	* @param propertyName  		- Name of the property
	* @param testRunner  		- Instance of test runner
	* @param log  				- Instance of log
	*
	* Revision History
	* Author			Date Modified			Comments
	* Kalesh Rajappan	09-Dec-2013				Initial Version
	*/	
	def readFile(String filePath,String propertyStepName,String propertyName, TestRunner testRunner,Logger log)	{
		
		def properties 	= testRunner.testCase.getTestStepByName(propertyStepName)
		String fileContents = new File(filePath).text		
		properties.setPropertyValue(propertyName,fileContents)
		//log.info "File read and content stored in property"
	}
		
}