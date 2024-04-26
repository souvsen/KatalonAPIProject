import static com.kms.katalon.core.checkpoint.CheckpointFactory.findCheckpoint 
import static com.kms.katalon.core.testcase.TestCaseFactory.findTestCase
import static com.kms.katalon.core.testdata.TestDataFactory.findTestData
import static com.kms.katalon.core.testobject.ObjectRepository.findTestObject
import static com.kms.katalon.core.testobject.ObjectRepository.findWindowsObject
import com.kms.katalon.core.checkpoint.Checkpoint as Checkpoint
import com.kms.katalon.core.cucumber.keyword.CucumberBuiltinKeywords as CucumberKW
import com.kms.katalon.core.mobile.keyword.MobileBuiltInKeywords as Mobile
import com.kms.katalon.core.model.FailureHandling as FailureHandling
import com.kms.katalon.core.testcase.TestCase as TestCase
import com.kms.katalon.core.testdata.TestData as TestData
import com.kms.katalon.core.testng.keyword.TestNGBuiltinKeywords as TestNGKW
import com.kms.katalon.core.testobject.TestObject as TestObject
import com.kms.katalon.core.webservice.keyword.WSBuiltInKeywords as WS
import com.kms.katalon.core.webui.keyword.WebUiBuiltInKeywords as WebUI
import com.kms.katalon.core.windows.keyword.WindowsBuiltinKeywords as Windows
import internal.GlobalVariable as GlobalVariable
import org.openqa.selenium.Keys as Keys
import org.apache.poi.ss.usermodel.DataFormatter
import org.apache.poi.ss.usermodel.Row
import org.apache.poi.ss.usermodel.Sheet
import org.apache.poi.ss.usermodel.Workbook
import org.apache.poi.ss.usermodel.WorkbookFactory




File f = new File(System.getProperty("user.dir")+"\\TestData\\testData.xlsx")
FileInputStream fis = new FileInputStream(f)
Workbook wb = WorkbookFactory.create(fis)
Sheet sheetName = wb.getSheet("login")

int totalRows = sheetName.getLastRowNum()
println(totalRows)
Row rowCells = sheetName.getRow(0)
int totalCols = rowCells.getLastCellNum()
println(totalCols)

DataFormatter format = new DataFormatter()
String [][]testData = new String [totalRows][totalCols]

for (int i = 1; i<=totalRows; i++) {
	for(int j = 0; j<totalCols; j++) {
		testData[i-1][j] = format.formatCellValue(sheetName.getRow(i).getCell(j))
		println(testData[i-1][j])
	}
}