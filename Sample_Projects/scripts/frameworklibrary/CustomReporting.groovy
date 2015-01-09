package scripts.frameworklibrary

import com.eviware.soapui.SoapUI
import com.eviware.soapui.model.support.TestRunListenerAdapter
import com.eviware.soapui.model.testsuite.TestRunContext
import com.eviware.soapui.model.testsuite.TestRunner
import org.apache.log4j.Logger

import javax.xml.parsers.DocumentBuilderFactory
import java.util.*
import java.text.SimpleDateFormat
import groovy.xml.XmlUtil

/*
* 
* @author : Kalesh Rajappan
* @date   : 12-Sep-2013
* Revision History
* Version	Author			Date Modified		Comments
* 1.0		Kalesh Rajappan	08-Nov-2013			Made the change to make the report path relative to soapUi project location
* 1.1		Kalesh Rajappan	09-Dec-2013		Updated the functions to create the HTML report based on the config property

* <ToDo> 
* 1. Make the reporting functions generic. Add more types of report
*/

class CustomReporting {

	/* 
	* Function Name 	: createReportFile
	* Description		: Creates the report file and store the file name in a property
	* @param reportBasePath  - Location for the report
	* @param testRunner  - Instance of test runner
	* @param context  - Instance of run context
	* @param log  - Instance of log
	* @return - None
	* Revision History
	* Author			Date Modified	Comments
	* Kalesh Rajappan	19-Sep-2013		Initial Version
	* Kalesh Rajappan	08-Nov-2013		Added a call to get the report file name through the function getReportFileName
	* Kalesh Rajappan	09-Dec-2013		Updated the function to create the HTML report based on the config property
	*/
	
   def createReportFile (String reportBasePath,TestRunner testRunner,TestRunContext context,Logger log) { 
   	  
		def NewReport = testRunner.testCase.testSuite.project.getPropertyValue( "GenerateNewReport" )
		//CreateHTMLReport value will be read from the config properties file and added to the project properties
		def GenerateHtmlReport = testRunner.testCase.testSuite.project.getPropertyValue( "CreateHTMLReport" )			
		def projectName = testRunner.testCase.testSuite.project.name
		def testSuiteName = testRunner.testCase.testSuite.name
		def testCaseName = testRunner.testCase.name		
				
		// Using the property at the project level to identify if the new report file needs to be created or not
		// Need to find out if there is any other option to implement it. Using the start up scripts in the soapUi test may complicate scripting
		if (NewReport == "true")
		{
			def fileName = getReportFileName (projectName,reportBasePath,context)
			
			//if CreateHTMLReport in properties file is false, don't generate the HTML file
			if (GenerateHtmlReport != "false")
			{
				createHtmlReport(fileName+".html",testSuiteName,testCaseName,testRunner, context, log)
			}
			testRunner.testCase.testSuite.project.setPropertyValue( "GenerateNewReport","false")
			
		}
		else
		{
			//if CreateHTMLReport in properties file is false, don't append to the HTML file
			if (GenerateHtmlReport != "false")
			{
				appendToHtmlReport("<table border='1' bordercolor='#000000' style='background-color:#FFFFFF' width='100%' cellpadding='3' cellspacing='3'><tr><td width='100%' colspan = '4' align = 'center'><b>Scenario : "+testSuiteName+" -- "+testCaseName+"<b></td></tr>",testRunner,context,log)
			}
		}

   }     
    /*
	* This methods prepares the the report file name
	* @param ReportName  - Name of the report
	* @param reportBasePath  - Location for the report
	* Revision History
	* Author			Date Modified			Comments
	* Kalesh Rajappan	08-Nov-2013				Initial Version
	*/
	def getReportFileName(String ReportName , String reportBasePath,TestRunContext context) {
		// get the current date and format it
		def date = new Date()
		def dateFormat = new java.text.SimpleDateFormat('yyyy_MM_dd_HHmmss')
		def currentDateTime = dateFormat.format(date)

		//if the value of Basepath is DEFAULT , set the report path relative to the project basepath 
		if (reportBasePath.toUpperCase() == "DEFAULT")
		{
		def groovyUtils = new com.eviware.soapui.support.GroovyUtils(context)
		reportBasePath = groovyUtils.projectPath + "\\Reports"
		}	
		
		//set the complete report file name
		def reportFilename = reportBasePath + "\\" + ReportName + "_" + currentDateTime
	}
	/*
	* This methods creates a new folder
	*/
	def createFolder(String testCaseName , String reportBasePath) {
	  def date = new Date()
	  def dateFormat = new java.text.SimpleDateFormat('MMddyyyy')
	  def shortDate = dateFormat.format(date)
	  def outputFolder = reportBasePath + testCaseName + "_" + shortDate
	  context.TestCaseReportFolder = outputFolder
	  createFolder = new File(outputFolder)
	  createFolder.mkdir()  
	}
	
    /*
	* This methods loops through the test steps and assertions and reports the status
	* @param testStepNames  - test steps which needs to be added in the report
	* @param reportBasePath  - Location for the report
	* @param testRunner  - Instance of test runner
	* @param context  - Instance of run context
	* @param log  - Instance of log
	* Revision History
	* Author			Date Modified			Comments
	* Kalesh Rajappan	12-Sep-2013				Initial Version
	* Kalesh Rajappan	09-Dec-2013				Removed the parameter testStepNames. made code change to read the step names from the current test case.
	*/
	def reportTestCaseLoopStatus(String propertyStepName,TestRunner testRunner,TestRunContext context,Logger log){
		def groovyUtils    =    new com.eviware.soapui.support.GroovyUtils( context )
		def ReportFile = context.getProperty("ReportFile")
		
		def rptMessage = ""
		def rptTestStatus = ""
		def testStepProperties = testRunner.testCase.testSteps[propertyStepName]
		def tcSlNo = testStepProperties.getPropertyValue("SlNo")
		def tcName = testStepProperties.getPropertyValue("TestCaseName")
		def assertionName
		def assertionStatus
		def assertionMessage
		//get the list of test steps
		def testSteps =  testRunner.testCase.getTestStepList()

		for( tstStep in testSteps )
		{
			//def tstStep = testRunner.testCase.testSteps[step] 	
			//def tstType =  tstStep.config.type
										
			if (tstStep instanceof com.eviware.soapui.model.testsuite.Assertable &&  !(tstStep.isDisabled()))
			{	
				def testStepName = tstStep.name
				//Loop through all the assertions
				for( assertion in tstStep.assertionList )
				{
				//log.info "Assertion [" + assertion.label + "] has status [" + assertion.status + "]"
					assertionName = assertion.label
					assertionStatus = assertion.status.toString()
					assertionMessage = ""
					rptMessage = rptMessage + "Step [ " + testStepName + " ] - Assertion [ " + assertionName + " ] " + assertionStatus +"<br>"
					
					//for each assertion , get the list of errors if any
					for( e in assertion.errors )
					{
						//errorMessage = errorMessage + "Assertion [ " +assertion.label + "] " + assertion.status + " "+ e.message
						assertionMessage = assertionMessage + " " + e.message
						rptMessage = rptMessage + "&nbsp;&nbsp;" + e.message +"<br>"
						context.TestCaseStatus = "FAIL"
					}
					//addAssertionRptXML(context.XmlRptDocument,context.XmlRptTestStepNode,assertionName,assertionStatus,assertionMessage,"","",context)
				}
			}
			//save the response. Need to add logic to write the request.
			//testRequest.requestContent will not work for request because the data will be shown as properties
			/*
			if (tstStep instanceof com.eviware.soapui.model.testsuite.SamplerTestStep && tstStep.testRequest.response != null )
			{
				//def holder = groovyUtils.getXmlHolder( step+"#Request" )
				//log.info holder.getXml()
				//testStepRequest = tstStep.testRequest.requestContent
				
				def testStepResponse = tstStep.testRequest.response.contentAsString      //response
				  String respFilename = "C:/Reports/Response/"+tcSlNo+"-"+tcName+"-response.xml"//getResponseFilename(testStepName)
				  def file = new PrintWriter (respFilename)
				  file.println(testStepResponse)
				  file.flush()
				  file.close()
			}
			*/
		}
		//Set the scenario description header for reporting.
		if (context.TestScenarioDesc == null)
		{
			context.TestScenarioDesc = "<html><head><title>Results - SoapUI Test Execution </title></head><body><table border='1' bordercolor='#000000' style='background-color:#FFFFFF' width='100%' cellpadding='3' cellspacing='3'><tr><td width='10%'><b>Sl #</b></td><td width='40%'><b>Test Case Name</b></td><td width='10%'><b>Status</b></td><td width='40%'><b>Comments</b></td></tr>"
		}
		rptMessage = context.TestStatusDesc + rptMessage
		
		if (context.TestCaseStatus == "FAIL")
		{
			rptTestStatus = "<font color='red'><b>" + context.TestCaseStatus + "</b></font>"
			context.TestScenarioStatus = "FAIL"
			//context.TestScenarioDesc = context.TestScenarioDesc + "<br>"+ tcSlNo + "|" + tcName + "|" + rptTestStatus + "|" + rptMessage
			context.TestScenarioDesc = context.TestScenarioDesc + "<tr><td width='10%'>" + tcSlNo + "</td><td width='40%'>"+tcName +"</td><td width='10%'>"+ rptTestStatus +"</td><td width='40%'>"+rptMessage+"</td></tr>"
		}
		else
		{
			rptTestStatus = "<font color='green'><b>" + context.TestCaseStatus + "</b></font>"
			//context.TestScenarioDesc = context.TestScenarioDesc + "<br>"+ tcSlNo + "|" + tcName + "|" + rptTestStatus
			context.TestScenarioDesc = context.TestScenarioDesc + "<tr><td width='10%'>" + tcSlNo + "</td><td width='40%'>"+tcName +"</td><td width='10%'>"+ rptTestStatus +"</td><td width='40%'>"+rptMessage+"</td></tr>"
		}
		
		//log.info rptMessage
		def htmlRptContent = "<tr><td width='10%'>" + tcSlNo + "</td><td width='40%'>"+tcName +"</td><td width='10%'>"+ rptTestStatus +"</td><td width='40%'>"+rptMessage+"</td></tr>"
		appendToHtmlReport(htmlRptContent,testRunner,context,log)

	}
	
		/*
	* This method creates the HTML report
	* @param fileName  - Name of the report file
	* @param testSuiteName  - Test suite name
	* @param testCaseName  - Test case name
	* @param testRunner  - Instance of test runner - currently not used. Added for future extension
	* @param context  - Instance of run context - currently not used. Added for future extension
	* @param log  - Instance of log - currently not used. Added for future extension
	* Revision History
	* Author			Date Modified			Comments
	* Kalesh Rajappan	09-Dec-2013				Initial Version
	*/
	def createHtmlReport(String fileName,String testSuiteName,String testCaseName,TestRunner testRunner,TestRunContext context,Logger log){
		//CreateHTMLReport value will be read from the config properties file and added to the project properties
		def GenerateHtmlReport = testRunner.testCase.testSuite.project.getPropertyValue( "CreateHTMLReport" )
		//if CreateHTMLReport in properties file is false, don't generate the HTML file
		if (GenerateHtmlReport != "false")
		{
			def ReportFile = new File (fileName)			
			ReportFile.write("<html><head><title>Results - SoapUI Test Execution </title></head><body><table border='1' bordercolor='#000000' style='background-color:#FFFFFF' width='100%' cellpadding='3' cellspacing='3'><tr><td align = 'center'><h2>SoapUI Test Execution Results</h2></td></tr></table>")
			ReportFile.append("<table border='1' bordercolor='#000000' style='background-color:#FFFFFF' width='100%' cellpadding='3' cellspacing='3'><tr><td width='10%'><b>Sl #</b></td><td width='40%'><b>Test Case Name</b></td><td width='10%'><b>Status</b></td><td width='40%'><b>Comments</b></td></tr>")
			ReportFile.append("<tr><td width='100%' colspan = '4' align = 'center'><b>Scenario : "+testSuiteName+" -- "+testCaseName+"<b></td></tr>")
			//context.setProperty("ReportFile", ReportFile)

			testRunner.testCase.testSuite.project.setPropertyValue( "HtmlReportFileName",fileName)
		}
	
	}

	/*
	* This method append contents to the HTML report
	* @param testStepNames  - test steps which needs to be added in the report
	* @param reportMessage  - message to be added to the report
	* @param testRunner  - Instance of test runner - currently not used. Added for future extension
	* @param context  - Instance of run context - currently not used. Added for future extension
	* @param log  - Instance of log - currently not used. Added for future extension
	* Revision History
	* Author			Date Modified			Comments
	* Kalesh Rajappan	12-Sep-2013				Initial Version
	*/
	def appendToHtmlReport(String reportMessage,TestRunner testRunner,TestRunContext context,Logger log){
		
		//CreateHTMLReport value will be read from the config properties file and added to the project properties
		def GenerateHtmlReport = testRunner.testCase.testSuite.project.getPropertyValue( "CreateHTMLReport" )
		//if CreateHTMLReport in properties file is false, don't generate the HTML file
		if (GenerateHtmlReport != "false")
		{
			def fileName = testRunner.testCase.testSuite.project.getPropertyValue( "HtmlReportFileName")
			def ReportFile	= new File(fileName)		
			ReportFile.append(reportMessage)
		}
	
	}
	/*
	* This method closes the HTML report
	* @param reportMessage  - message to be added to the report
	* @param testRunner  - Instance of test runner - currently not used. Added for future extension
	* @param context  - Instance of run context - currently not used. Added for future extension
	* @param log  - Instance of log - currently not used. Added for future extension
	* Revision History
	* Author			Date Modified			Comments
	* Kalesh Rajappan	12-Sep-2013				Initial Version
	*/
	def closeReportFile(String reportMessage,TestRunner testRunner,TestRunContext context,Logger log){
		
		//CreateHTMLReport value will be read from the config properties file and added to the project properties
		def GenerateHtmlReport = testRunner.testCase.testSuite.project.getPropertyValue( "CreateHTMLReport" )
		//if CreateHTMLReport in properties file is false, don't generate the HTML file
		if (GenerateHtmlReport != "false")
		{
			def fileName = testRunner.testCase.testSuite.project.getPropertyValue( "HtmlReportFileName")
			def ReportFile	= new File(fileName)
			ReportFile.append("</table></body></html>")
			context.TestScenarioDesc = context.TestScenarioDesc + "</table></body></html>"
		}
	
	}
	/*
	* This method is used to create XML reports. Currently XML reports are not used
	*/
	def addTestSuiteRptXML(document, XmlRptRoot, testSuiteName, startTime,TestRunContext context){
		def XmlRptTestSuiteNode = document.createElement('testSuite')
		XmlRptTestSuiteNode.setAttribute('name', testSuiteName)
		XmlRptTestSuiteNode.setAttribute('startTime', startTime)
		XmlRptRoot.appendChild(XmlRptTestSuiteNode)
		context.XmlRptTestSuiteNode = XmlRptTestSuiteNode
	   // return XmlRptTestSuiteNode
	}

	def addTestCaseRptXML(document,XmlRptTestSuiteNode , testCaseName, startTime,TestRunContext context){
		def XmlRptTestCaseNode = document.createElement('testCase')
		XmlRptTestCaseNode.setAttribute('name', testCaseName)
		XmlRptTestCaseNode.setAttribute('startTime', startTime)
		XmlRptTestSuiteNode.appendChild(XmlRptTestCaseNode)
		context.XmlRptTestCaseNode = XmlRptTestCaseNode
		//return XmlRptTestCaseNode
	}

	def addTestStepRptXML(document,XmlRptTestCaseNode , testStepName, startTime,TestRunContext context){
		def XmlRptTestStepNode = document.createElement('testStep')
		XmlRptTestStepNode.setAttribute('name', testStepName)
		XmlRptTestStepNode.setAttribute('startTime', startTime)
		XmlRptTestCaseNode.appendChild(XmlRptTestStepNode)
		context.XmlRptTestStepNode= XmlRptTestStepNode
		//return XmlRptTestStepNode
	}

	def addAssertionRptXML(document,XmlRptTestStepNode , assertionName, status, message, expected, actual,TestRunContext context){
		def XmlRptAssertionNode = document.createElement('assertion')
		XmlRptAssertionNode.setAttribute('name', assertionName)
		XmlRptAssertionNode.setAttribute('status', status)
		//XmlRptAssertionNode.setAttribute('message', startTime)
		XmlRptTestStepNode.appendChild(XmlRptAssertionNode)

		def messageNode = document.createElement('message')
			messageNode.appendChild(document.createTextNode(message))
		   XmlRptAssertionNode.appendChild(messageNode)

		   def expValueNode = document.createElement('expected')
			expValueNode.appendChild(document.createTextNode(expected))
		   XmlRptAssertionNode.appendChild(expValueNode)

		   def actualValueNode = document.createElement('actual')
			actualValueNode.appendChild(document.createTextNode(actual))
		   XmlRptAssertionNode.appendChild(actualValueNode)
	   
		//return XmlRptAssertionNode
	}

	def updateTestSuiteRptStatusXML(XmlRptTestSuiteNode , status,TestRunContext context){
		XmlRptTestSuiteNode.setAttribute('status', status)   
	}

	def updateTestCaseRptStatusXML(XmlRptTestCaseNode , status,TestRunContext context){
		XmlRptTestCaseNode.setAttribute('status', status)   
	}

	def updateTestStepRptStatusXML(XmlRptTestStepNode , status,TestRunContext context){
		XmlRptTestStepNode.setAttribute('status', status)   
	}


	def createXMLReportDocument(TestRunContext context)
	{
		def builder  = DocumentBuilderFactory.newInstance().newDocumentBuilder()
		def document = builder.newDocument()
		def XmlRptRoot     = document.createElement('testResults')
		document.appendChild(XmlRptRoot)
		
		context.XmlRptDocument = document
		context.XmlRptRoot = XmlRptRoot
	}

}

