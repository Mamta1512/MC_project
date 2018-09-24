package com.utils;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.w3c.dom.Document;

public class ExcelUtils {
	public ExcelUtils() {
		// TODO Auto-generated constructor stub
	}

	public static void main(String[] args) throws IOException {
		ExcelUtils.parseExcel();
	}

	public static List<Screen> parseExcel() throws IOException {
		List<Screen> screens = new ArrayList<Screen>();

		DocumentBuilderFactory dbFactory = DocumentBuilderFactory.newInstance();
		DocumentBuilder dBuilder = null;
		try {
			dBuilder = dbFactory.newDocumentBuilder();
		} catch (ParserConfigurationException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		Document doc = dBuilder.newDocument();

		FileInputStream file = new FileInputStream(new File("ObjectRepository.xls"));

//Create Workbook instance holding reference to .xlsx file
		HSSFWorkbook workbook = new HSSFWorkbook(file);

//Get first/desired sheet from the workbook
		HSSFSheet sheet = workbook.getSheetAt(0);
		Screen screen = new Screen();
//Iterate through each rows one by one
		Iterator<Row> rowIterator = sheet.iterator();
		while (rowIterator.hasNext()) {
			Row row = rowIterator.next();

			if (row.getRowNum() == 0)
				continue;
//For each row, iterate through all the columns
			Iterator<Cell> cellIterator = row.cellIterator();
			List<Cell> cells = new ArrayList<>();
			cellIterator.forEachRemaining(cells::add);

			for (int i = 0; i < cells.size(); i++) {
				if (cells.get(0).getStringCellValue().toString() != screen.Name) {
					screen = new Screen();
					screen.Name = cells.get(0).getStringCellValue().toString();
					screens.add(screen);
				} else
					screen.controls.put(cells.get(1).getStringCellValue().toString(),
							cells.get(2).getStringCellValue().toString());
			}

		}
		file.close();

		return screens;
	}
}
