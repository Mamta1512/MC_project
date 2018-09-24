package com.excelReader.ReadExcel;

import java.io.File;
import java.io.FileInputStream;
import java.util.Iterator;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.w3c.dom.Document;
import org.w3c.dom.Element;

public class demo {

	public static void main(String[] args) {
		try {
			DocumentBuilderFactory dbFactory = DocumentBuilderFactory.newInstance();
			DocumentBuilder dBuilder = dbFactory.newDocumentBuilder();
			Document doc = dBuilder.newDocument();

			FileInputStream file = new FileInputStream(new File("ObjectRepo.xls"));

			Element rootElement = doc.createElement("Screens");

			// Create Workbook instance holding reference to .xlsx file
			HSSFWorkbook workbook = new HSSFWorkbook(file);

			// Get first/desired sheet from the workbook
			HSSFSheet sheet = workbook.getSheetAt(0);

			// Iterate through each rows one by one
			Iterator<Row> rowIterator = sheet.iterator();

			String screenName = null;
			Element screenNameElement = null;
			Element control = null;

			try {
				while (rowIterator.hasNext()) {
					Row row = rowIterator.next();

					if (row.getRowNum() == 0)
						continue;
					// For each row, iterate through all the columns
					Iterator<Cell> cellIterator = row.cellIterator();

					while (cellIterator.hasNext()) {
						Cell cell = cellIterator.next();

						if (cell.getColumnIndex() == 0) {
							if (screenName != cell.getStringCellValue()) {

								screenName = cell.getStringCellValue();
								screenNameElement = doc.createElement(screenName);

								if (screenNameElement != null) {
									rootElement.appendChild(screenNameElement);
								}
							}
						}
						if (cell.getColumnIndex() == 1) {
							if (cell.getStringCellValue().toString() == "")

							{
								control = null;
								continue;
							}
							control = doc.createElement(cell.getStringCellValue());
							screenNameElement.appendChild(control).toString();
						}
						if (cell.getColumnIndex() == 2 && control != null) {
							Element xpath = doc.createElement("XPath");
							xpath.appendChild(doc.createTextNode(cell.getStringCellValue()));
							control.appendChild(xpath).toString();
						}

					}

				}

				doc.appendChild(rootElement);
			} catch (Exception e) {
				e.printStackTrace();

			}
			file.close();

			// write the content into xml file
			TransformerFactory transformerFactory = TransformerFactory.newInstance();
			Transformer transformer = transformerFactory.newTransformer();
			DOMSource source = new DOMSource(doc);
			StreamResult result = new StreamResult(new File("D:\\Repo.xml"));
			transformer.transform(source, result);

			// Output to console for testing
			StreamResult consoleResult = new StreamResult(System.out);
			transformer.transform(source, consoleResult);
		} catch (Exception e) {
			e.printStackTrace();
		}

	}

}
