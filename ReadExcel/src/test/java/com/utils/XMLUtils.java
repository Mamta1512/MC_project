package com.utils;

import java.io.File;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;

import org.w3c.dom.Document;
import org.w3c.dom.Element;

public class XMLUtils {

	public static void main(String[] args) throws Exception {
		List<Screen> screens = ExcelUtils.parseExcel();
		WriteObjectMap(screens, "D:\\repo.xml");
	}

	public static void WriteObjectMap(List<Screen> screens, String XMLFile) throws Exception {

		String ScreenName = null;
		Element controlName = null;

		DocumentBuilderFactory dbFactory = DocumentBuilderFactory.newInstance();
		DocumentBuilder dBuilder = dbFactory.newDocumentBuilder();
		Document doc = dBuilder.newDocument();

		Element rootElement = doc.createElement("Screens");
		doc.appendChild(rootElement);

		for (int i = 0; i < screens.size(); i++) {

			ScreenName = screens.get(i).Name.toString();
			Element ScreenElement = doc.createElement(ScreenName);
			rootElement.appendChild(ScreenElement);
			HashMap<String, String> elements = screens.get(i).controls;

			for (Map.Entry<String, String> element : elements.entrySet()) {

//				if (element.getKey() == "")

//				{
//
//					// Element xpath = doc.createElement("XPath");
//					// xpath.appendChild(doc.createTextNode(element.getValue()));
//					// ScreenElement.appendChild(xpath);
//					continue;
//
//				} else 
				if (element.getKey() != "") {
					Element logicalName = doc.createElement(element.getKey());
					ScreenElement.appendChild(logicalName);
					Element xpath = doc.createElement("XPath");
					xpath.appendChild(doc.createTextNode(element.getValue()));
					logicalName.appendChild(xpath);
				}

			}

		}

		// write the content into xml file
		TransformerFactory transformerFactory = TransformerFactory.newInstance();
		Transformer transformer = transformerFactory.newTransformer();
		DOMSource source = new DOMSource(doc);
		StreamResult result = new StreamResult(new File("D:\\Repo.xml"));
		transformer.transform(source, result);

		// Output to console for testing
		StreamResult consoleResult = new StreamResult(System.out);
		transformer.transform(source, consoleResult);
	}
}
