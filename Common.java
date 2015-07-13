package odesk.scraper.utilities;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.DecimalFormat;
import org.apache.pdfbox.cos.COSDocument;
import org.apache.pdfbox.pdfparser.PDFParser;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.util.PDFTextStripper;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.firefox.FirefoxProfile;
import org.openqa.selenium.phantomjs.PhantomJSDriver;
import org.openqa.selenium.phantomjs.PhantomJSDriverService;
import org.openqa.selenium.remote.DesiredCapabilities;

public class Common {
	static WebDriver driver;

	public static FirefoxProfile firefoxProfile() {

		FirefoxProfile firefoxProfile = new FirefoxProfile();
		firefoxProfile.setPreference("browser.download.folderList", 2);
		firefoxProfile.setPreference(
				"browser.download.manager.showWhenStarting", false);
		// firefoxProfile.setPreference("browser.download.dir", downloadPath);
		firefoxProfile.setPreference("browser.download.dir",
				"D:\\WebDriverDownloadFolder");
		firefoxProfile.setPreference("browser.helperApps.neverAsk.saveToDisk",
				"application/pdf");

		firefoxProfile.setPreference("pdfjs.disabled", true);

		// Use this to disable Acrobat plugin for previewing PDFs in Firefox (if
		// you have Adobe reader installed on your computer)
		firefoxProfile.setPreference("plugin.scan.Acrobat", "99.0");
		firefoxProfile.setPreference("plugin.scan.plid.all", false);

		return firefoxProfile;
	}

	public WebDriver getBrowser(WebDriver driver, String browser) {
		switch (browser) {
		case "Firefox":
			try {
				driver = new FirefoxDriver(firefoxProfile());
			} catch (Exception e2) {
				// TODO Auto-generated catch block
				System.out.println("Unable to using Firefox Profile.");
				driver = new FirefoxDriver();
			}
			break;

		case "Chrome":
			System.setProperty("webdriver.chrome.driver",
					"D:/oDesk/chromedriver.exe");
			driver = new ChromeDriver();
			break;

		case "Headless":
			// prepare capabilities
			DesiredCapabilities DesireCaps = new DesiredCapabilities();
			DesireCaps.setCapability("takesScreenshot", true);
			DesireCaps
					.setCapability(
							PhantomJSDriverService.PHANTOMJS_EXECUTABLE_PATH_PROPERTY,
							"D:\\DevTools\\phantomjs-2.0.0-windows\\bin\\phantomjs.exe");
			driver = new PhantomJSDriver(DesireCaps);
			break;

		}
		return driver;
	}

	public void writeExcel(String filePath, String fileName, String sheetName,
			String[] dataToWrite) throws IOException {

		// Create a object of File class to open xlsx file
		File file = new File(filePath + "\\" + fileName);

		// Create an object of FileInputStream class to read excel file
		FileInputStream inputStream = new FileInputStream(file);
		Workbook writableWorkbook = null;

		// Find the file extension by spliting file name in substing and getting
		// only extension name
		String fileExtensionName = fileName.substring(fileName.indexOf("."));

		// Check condition if the file is xlsx file
		if (fileExtensionName.equals(".xlsx")) {
			// If it is xlsx file then create object of XSSFWorkbook class
			writableWorkbook = new XSSFWorkbook(inputStream);
		}

		// Check condition if the file is xls file
		else if (fileExtensionName.equals(".xls")) {
			// If it is xls file then create object of XSSFWorkbook class
			writableWorkbook = new HSSFWorkbook(inputStream);
		}

		// Read excel sheet by sheet name
		Sheet sheet = writableWorkbook.getSheet(sheetName);

		// Get the current count of rows in excel file
		int rowCount = sheet.getLastRowNum() - sheet.getFirstRowNum();

		// Get the first row from the sheet
		Row row = sheet.getRow(0);

		// Create a new row and append it at last of sheet
		Row newRow = sheet.createRow(rowCount + 1);

		// Create a loop over the cell of newly created Row
		for (int j = 0; j < row.getLastCellNum(); j++) {
			// Fill data in row
			Cell cell = newRow.createCell(j);
			cell.setCellValue(dataToWrite[j]);
		}

		// Close input stream
		inputStream.close();

		// Create an object of FileOutputStream class to create write data in
		// excel file
		FileOutputStream outputStream = new FileOutputStream(file);

		// write data in the excel file
		writableWorkbook.write(outputStream);

		// close output stream
		outputStream.close();
	}

	public void writeExcel(String filePath, String fileName, String sheetName,
			String[][] dataToWrite, int noOfRows) throws IOException {

		// Create a object of File class to open xlsx file
		File file = new File(filePath + "\\" + fileName);

		// Create an object of FileInputStream class to read excel file
		FileInputStream inputStream = new FileInputStream(file);
		Workbook writableWorkbook = null;

		// Find the file extension by spliting file name in substing and getting
		// only extension name
		String fileExtensionName = fileName.substring(fileName.indexOf("."));

		// Check condition if the file is xlsx file
		if (fileExtensionName.equals(".xlsx")) {
			// If it is xlsx file then create object of XSSFWorkbook class
			writableWorkbook = new XSSFWorkbook(inputStream);
		}

		// Check condition if the file is xls file
		else if (fileExtensionName.equals(".xls")) {
			// If it is xls file then create object of XSSFWorkbook class
			writableWorkbook = new HSSFWorkbook(inputStream);
		}

		// Read excel sheet by sheet name
		Sheet sheet = writableWorkbook.getSheet(sheetName);

		// Get the current count of rows in excel file
		int rowCount = sheet.getLastRowNum() - sheet.getFirstRowNum();

		// Get the first row from the sheet
		Row row = sheet.getRow(0);

		// Create a loop over the cell of newly created Row
		for (int i = 0; i < noOfRows; i++) {
			// Create a new row and append it at last of sheet
			Row newRow = sheet.createRow(rowCount + 1 + i);
			for (int j = 0; j < row.getLastCellNum(); j++) {
				// Fill data in row
				Cell cell = newRow.createCell(j);
				cell.setCellValue(dataToWrite[i][j]);
			}
		}

		// Close input stream
		inputStream.close();

		// Create an object of FileOutputStream class to create write data in
		// excel file
		FileOutputStream outputStream = new FileOutputStream(file);

		// write data in the excel file
		writableWorkbook.write(outputStream);

		// close output stream
		outputStream.close();
	}

	public void getNumberOfCells(String filePath, String fileName,
			String sheetName, String[] dataToWrite) throws IOException {

		// Create a object of File class to open xlsx file
		File file = new File(filePath + "\\" + fileName);

		// Create an object of FileInputStream class to read excel file
		FileInputStream inputStream = new FileInputStream(file);
		Workbook writableWorkbook = null;

		// Find the file extension by spliting file name in substing and getting
		// only extension name
		String fileExtensionName = fileName.substring(fileName.indexOf("."));

		// Check condition if the file is xlsx file
		if (fileExtensionName.equals(".xlsx")) {
			// If it is xlsx file then create object of XSSFWorkbook class
			writableWorkbook = new XSSFWorkbook(inputStream);
		}

		// Check condition if the file is xls file
		else if (fileExtensionName.equals(".xls")) {
			// If it is xls file then create object of XSSFWorkbook class
			writableWorkbook = new HSSFWorkbook(inputStream);
		}

		// Read excel sheet by sheet name
		Sheet sheet = writableWorkbook.getSheet(sheetName);

		// Get the current count of rows in excel file
		int rowCount = sheet.getLastRowNum() - sheet.getFirstRowNum();

		// Get the first row from the sheet
		Row row = sheet.getRow(0);

		// Create a new row and append it at last of sheet
		Row newRow = sheet.createRow(rowCount + 1);

		// Create a loop over the cell of newly created Row
		for (int j = 0; j < row.getLastCellNum(); j++) {
			// Fill data in row
			Cell cell = newRow.createCell(j);
			cell.setCellValue(dataToWrite[j]);
		}

		// Close input stream
		inputStream.close();

		// Create an object of FileOutputStream class to create write data in
		// excel file
		FileOutputStream outputStream = new FileOutputStream(file);

		// write data in the excel file
		writableWorkbook.write(outputStream);

		// close output stream
		outputStream.close();
	}

	public void readExcel(String filePath, String fileName, String sheetName)
			throws IOException {

		// Create a object of File class to open xlsx file
		File file = new File(filePath + "\\" + fileName);

		// Create an object of FileInputStream class to read excel file
		FileInputStream inputStream = new FileInputStream(file);
		Workbook readableWorkbook = null;

		// Find the file extension by spliting file name in substring and
		// getting only extension name
		String fileExtensionName = fileName.substring(fileName.indexOf("."));

		// Check condition if the file is xlsx file
		if (fileExtensionName.equals(".xlsx")) {
			// If it is xlsx file then create object of XSSFWorkbook class
			readableWorkbook = new XSSFWorkbook(inputStream);
		}

		// Check condition if the file is xls file
		else if (fileExtensionName.equals(".xls")) {
			// If it is xls file then create object of XSSFWorkbook class
			readableWorkbook = new HSSFWorkbook(inputStream);
		}

		// Read sheet inside the workbook by its name
		Sheet guru99Sheet = readableWorkbook.getSheet(sheetName);

		// Find number of rows in excel file
		int rowCount = guru99Sheet.getLastRowNum()
				- guru99Sheet.getFirstRowNum();

		// Create a loop over all the rows of excel file to read it
		for (int i = 0; i < rowCount + 1; i++) {
			Row row = guru99Sheet.getRow(i);

			// Create a loop to print cell values in a row
			for (int j = 0; j < row.getLastCellNum(); j++) {
				// Print excel data in console
				System.out.println(row.getCell(j).getStringCellValue());
			}
			System.out.println();
		}

	}

	public String readValueFromExcel(String filePath, String fileName,
			String sheetName, int indexOfRow, int indexOfColumn)
			throws IOException {
		// Format Double
		DecimalFormat df = new DecimalFormat("#");

		// Create a object of File class to open xlsx file
		File file = new File(filePath + "\\" + fileName);

		// Create an object of FileInputStream class to read excel file
		FileInputStream inputStream = new FileInputStream(file);
		Workbook readableWorkbook = null;

		// Find the file extension by spliting file name in substring and
		// getting only extension name
		String fileExtensionName = fileName.substring(fileName.indexOf("."));

		// Check condition if the file is xlsx file
		if (fileExtensionName.equals(".xlsx")) {
			// If it is xlsx file then create object of XSSFWorkbook class
			readableWorkbook = new XSSFWorkbook(inputStream);
		}

		// Check condition if the file is xls file
		else if (fileExtensionName.equals(".xls")) {
			// If it is xls file then create object of XSSFWorkbook class
			readableWorkbook = new HSSFWorkbook(inputStream);
		}

		// Read sheet inside the workbook by its name
		Sheet guru99Sheet = readableWorkbook.getSheet(sheetName);

		// Find number of rows in excel file
		int rowCount = guru99Sheet.getLastRowNum()
				- guru99Sheet.getFirstRowNum();

		// Get the row
		if (indexOfRow <= rowCount) {
			Object value;
			Row row = guru99Sheet.getRow(indexOfRow);
			switch (row.getCell(indexOfColumn).getCellType()) {
			case Cell.CELL_TYPE_STRING:
				value = row.getCell(indexOfColumn).getStringCellValue();
				break;
			case Cell.CELL_TYPE_NUMERIC:
				value = df.format(row.getCell(indexOfColumn)
						.getNumericCellValue());
				break;
			default:
				throw new RuntimeException(
						"There is no support for this type of cell");
			}
			return value.toString();
		} else {
			return "No Value";
		}
	}

	public void readValuesFromAllCells(String filePath, String fileName,
			String sheetName, int indexOfColumn) throws IOException {

		// Create a object of File class to open xlsx file
		File file = new File(filePath + "\\" + fileName);

		// Create an object of FileInputStream class to read excel file
		FileInputStream inputStream = new FileInputStream(file);
		Workbook readableWorkbook = null;

		// Find the file extension by spliting file name in substring and
		// getting only extension name
		String fileExtensionName = fileName.substring(fileName.indexOf("."));

		// Check condition if the file is xlsx file
		if (fileExtensionName.equals(".xlsx")) {
			// If it is xlsx file then create object of XSSFWorkbook class
			readableWorkbook = new XSSFWorkbook(inputStream);
		}

		// Check condition if the file is xls file
		else if (fileExtensionName.equals(".xls")) {
			// If it is xls file then create object of XSSFWorkbook class
			readableWorkbook = new HSSFWorkbook(inputStream);
		}

		// Read sheet inside the workbook by its name
		Sheet guru99Sheet = readableWorkbook.getSheet(sheetName);

		// Find number of rows in excel file
		int rowCount = guru99Sheet.getLastRowNum()
				- guru99Sheet.getFirstRowNum();

		// Create a loop over all the rows of excel file to read it
		for (int i = 0; i < rowCount; i++) {
			Row row = guru99Sheet.getRow(i);

			// Create a loop to print cell values in a row
			// for (int j = 0; j < row.getLastCellNum(); j++) {
			// Print excel data in console
			System.out.print(row.getCell(indexOfColumn).getStringCellValue());
		}
	}

	public String extractTextFromPDF(String Folder, String fileName) {
		String value = "";
		PDFParser parser;
		PDFTextStripper pdfStripper = null;
		PDDocument pdDoc = null;
		COSDocument cosDoc = null;

		// File file = new File("D:/Downloads/Documents/StatementSearch.pdf");
		File file = new File(Folder + "\\" + fileName);
		if (!file.isFile()) {
			System.err.println("Folder: " + Folder + "\n\tFile: " + fileName
					+ " does not exist.");
			return null;
		}
		try {
			parser = new PDFParser(new FileInputStream(file));
		} catch (IOException e) {
			System.err.println("Unable to open PDF Parser. " + e.getMessage());
			return null;
		}

		try {
			parser.parse();
			cosDoc = parser.getDocument();
			pdfStripper = new PDFTextStripper();
			pdDoc = new PDDocument(cosDoc);
			pdfStripper.setStartPage(1);
			pdfStripper.setEndPage(1);
			// String parsedText = pdfStripper.getText(pdDoc);
			// System.out.println(parsedText);
			value = pdfStripper.getText(pdDoc);
		} catch (Exception e) {
			System.err.println("An exception occured in parsing the PDF: "
					+ e.getMessage());

		} finally {
			try {
				if (cosDoc != null)
					cosDoc.close();
				if (pdDoc != null)
					pdDoc.close();
			} catch (Exception e2) {
				e2.printStackTrace();
			}
		}
		return value;
	}

	public boolean isFileDownloaded(String downloadPath, String fileName) {
		boolean flag = false;
		File dir = new File(downloadPath);
		File[] dir_contents = dir.listFiles();

		for (int i = 0; i < dir_contents.length; i++) {
			if (dir_contents[i].getName().equals(fileName))
				return flag = true;
		}

		return flag;
	}

	// Main function is calling readExcel/writeExcel function to read data from
	// excel file
	public static void main(String[] args) {

		/***********************************************************************
		 * Writing excel file example
		 **********************************************************************/
		/*
		 * 
		 * // Create an array with the data in the same order in which you
		 * expect // to be filled in excel file String[] valueToWrite = {
		 * "Mr. E", "Noida", "3", "4", "5", "6", "7", "8", "9", "10", "11"};
		 * String[] valueToWrite1 = { "Mr. E", "Noida", "3", "4", "5", "6", "7",
		 * "8", "9", "10", "12"}; // Create an object of current class Common
		 * objWriteExcelFile = new Common();
		 * 
		 * // Write the file using file name , sheet name and the data to be
		 * filled objWriteExcelFile.writeExcel(System.getProperty("user.dir") +
		 * "\\temp", "ExportExcel.xls", "Sheet1", valueToWrite);
		 * 
		 * objWriteExcelFile.writeExcel(System.getProperty("user.dir") +
		 * "\\temp", "ExportExcel.xls", "Sheet1", valueToWrite1);
		 * 
		 * 
		 * /*********************************************************************
		 * ** Reading excel file example
		 * ********************************************************************
		 */
		/*
		 * 
		 * // Create a object of ReadGuru99ExcelFile class Common
		 * objReadExcelFile = new Common();
		 * 
		 * // Prepare the path of excel file String filePath =
		 * System.getProperty("user.dir") + "\\temp";
		 * 
		 * // Call read file method of the class to read data
		 * objReadExcelFile.readExcel(filePath, "ExportExcel.xls", "Sheet1");
		 */

		/***********************************************************************
		 * Get no of column in excel file example
		 **********************************************************************/
		/*
		 * // Create a object of File class to open xlsx file String fileName =
		 * "temp\\ExportExcel.xls";
		 * 
		 * File file = new File(fileName);
		 * 
		 * // Create an object of FileInputStream class to read excel file
		 * FileInputStream inputStream = new FileInputStream(file); Workbook
		 * workbook = null;
		 * 
		 * // Find the file extension by spliting file name in substing and
		 * getting // only extension name String fileExtensionName =
		 * fileName.substring(fileName.indexOf("."));
		 * 
		 * // Check condition if the file is xlsx file if
		 * (fileExtensionName.equals(".xlsx")) { // If it is xlsx file then
		 * create object of XSSFWorkbook class workbook = new
		 * XSSFWorkbook(inputStream); }
		 * 
		 * // Check condition if the file is xls file else if
		 * (fileExtensionName.equals(".xls")) { // If it is xls file then create
		 * object of XSSFWorkbook class workbook = new
		 * HSSFWorkbook(inputStream); } Sheet sheet =
		 * workbook.getSheet("Sheet1");
		 * 
		 * int noOfColumns = sheet.getRow(0).getLastCellNum();
		 * System.out.println("number of columns " + noOfColumns);
		 */

		/***********************************************************************
		 * Writing excel file example using multi dimension array
		 **********************************************************************/
		/*
		 * // Create an array with the data in the same order in which you
		 * expect // to be filled in excel file String[] valueToWrite = {
		 * "Mr. E", "Noida", "3", "4", "5", "6", "7", "8", "9", "10"};
		 * 
		 * String[][] valueToAdd = { { "Row1", "Row1", "Row1", "Row1", "Row1",
		 * "Row1", "Row1", "Row1", "Row1", "Row1" }, { "Row2", "Row2", "Row2",
		 * "Row2", "Row2", "Row2", "Row1", "Row1", "Row1", "Row1" }, { "Row3",
		 * "Row3", "Row3", "Row3", "Row2", "Row2", "Row1", "Row1", "Row1",
		 * "Row1" }, };
		 * 
		 * 
		 * // Create an object of current class Common objWriteExcelFile = new
		 * Common();
		 * 
		 * String folderName = "D:\\oDesk\\Scraper\\"; String fileName =
		 * "ExportExcel.xls";
		 * 
		 * // Write the file using file name , sheet name and the data to be
		 * filled objWriteExcelFile.writeExcel(folderName, fileName, "Sheet1",
		 * valueToAdd, 3);
		 */

		/***********************************************************************
		 * Reading excel file by index of column example
		 **********************************************************************/
		/*
		 * // Create a object of ReadGuru99ExcelFile class Common
		 * objReadExcelFile = new Common();
		 * 
		 * // Prepare the path of excel file String folderName =
		 * "D:\\oDesk\\Scraper\\"; String fileName = "ExportExcel.xls";
		 * 
		 * // Call read file method of the class to read data /*
		 * objReadExcelFile.readValuesFromAllCells(folderName, fileName,
		 * "Sheet1", 2);
		 */
		/*
		 * long startTime = System.currentTimeMillis(); String BBL =
		 * objReadExcelFile.readValueFromExcel(folderName, fileName, "Sheet1",
		 * 1, 1);
		 * 
		 * String blockNo = BBL.substring(2, 6); String loNo = BBL.substring(7);
		 * System.out.println("BBL  : " + BBL); System.out.println("BLOCK: " +
		 * blockNo); System.out.println("LO: " + loNo);
		 * 
		 * Date dNow = new Date(); SimpleDateFormat ft = new
		 * SimpleDateFormat("yyyy.MM.dd'_'hh_mm_ss_S");
		 * System.out.println("Start Time: " + ft.format(dNow));
		 * System.out.println(dNow); long endTime = System.currentTimeMillis();
		 * System.out.println("That took " + (endTime - startTime) +
		 * " milliseconds");
		 */

		/***********************************************************************
		 * Extract the text from PDF file
		 **********************************************************************/

		Common obj = new Common();
		String folder = "C:/Users/hmy1hc/Downloads/sample";
		String file = "1959_table_71.pdf";
		System.out.println(obj.isFileDownloaded(folder, file));
		String extractedText = obj.extractTextFromPDF(folder, file);

		System.out.println(extractedText);

	}

	public String extractAddress(String text) {
		String extractedText = "";
		String[] sentence = text.split("\n");
		int start = 0;
		int end = 0;
		// System.out.println("Length of Sentence Array: " + sentence.length);
		for (int index = 0; index < sentence.length; index++) {
			if (sentence[index].startsWith("Mailing address:")) {
				// System.out.println("Get index of Mailing Address word: " +
				// index);
				start = index + 1;
			}
			if (sentence[index].startsWith("Owner name:")) {
				// System.out.println("Get index of Owner name: " + index);
				end = index;
			}
		}
		while (start < end) {
			extractedText = extractedText + sentence[start].trim();
			extractedText = extractedText + "|";
			start++;
		}
		return extractedText;
	}

	public void loadingIcon(WebDriver driver, String Xpath){
    	long end = System.currentTimeMillis() + 60000;
        while (System.currentTimeMillis() < end) {
        	pause(1000);
            WebElement resultsDiv = driver.findElement(By.xpath(Xpath));
            if (!resultsDiv.isDisplayed()){     
              break;
            }
        }
    }
	
	public static void pause(final int iTimeInMillis) {
		try {
			Thread.sleep(iTimeInMillis);
		} catch (InterruptedException ex) {
			System.out.println(ex.getMessage());
		}
	}
}
