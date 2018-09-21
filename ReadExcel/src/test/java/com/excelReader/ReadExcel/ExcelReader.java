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
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.w3c.dom.Attr;
import org.w3c.dom.Document;
import org.w3c.dom.Element;


class CreateXMLFile
{
	public void createXML()
	{
		try {
			DocumentBuilderFactory dbFactory = DocumentBuilderFactory.newInstance();
	         DocumentBuilder dBuilder = dbFactory.newDocumentBuilder();
	         Document doc = dBuilder.newDocument();
			
			
			 Element rootElement = doc.createElement("cars");
	         doc.appendChild(rootElement);

	         // supercars element
	         Element supercar = doc.createElement("supercars");
	         rootElement.appendChild(supercar);

	         // setting attribute to element
	         Attr attr = doc.createAttribute("company");
	         attr.setValue("Ferrari");
	         supercar.setAttributeNode(attr);

	         // carname element
	         Element carname = doc.createElement("carname");
	         Attr attrType = doc.createAttribute("type");
	         attrType.setValue("formula one");
	         carname.setAttributeNode(attrType);
	         carname.appendChild(doc.createTextNode("Ferrari 101"));
	         supercar.appendChild(carname);

	         Element carname1 = doc.createElement("carname");
	         Attr attrType1 = doc.createAttribute("type");
	         attrType1.setValue("sports");
	         carname1.setAttributeNode(attrType1);
	         carname1.appendChild(doc.createTextNode("Ferrari 202"));
	         supercar.appendChild(carname1);

	         // write the content into xml file
	         TransformerFactory transformerFactory = TransformerFactory.newInstance();
	         Transformer transformer = transformerFactory.newTransformer();
	         DOMSource source = new DOMSource(doc);
	         StreamResult result = new StreamResult(new File("C:\\Users\\mamta_alane\\Documents\\cars.xml"));
	         transformer.transform(source, result);
	         
	         // Output to console for testing
	         StreamResult consoleResult = new StreamResult(System.out);
	         transformer.transform(source, consoleResult);

	      } catch (Exception e) {
	         e.printStackTrace();
	      }

	}
}





public class ExcelReader {

	public static void main(String[] args) {
		// TODO Auto-generated method stub

		try
        {
			DocumentBuilderFactory dbFactory = DocumentBuilderFactory.newInstance();
			         DocumentBuilder dBuilder = dbFactory.newDocumentBuilder();
			         Document doc = dBuilder.newDocument();
			         
			         
			         Element rootElement = doc.createElement("Screens");
			         doc.appendChild(rootElement);
			
			
            FileInputStream file = new FileInputStream(new File("ObjectRepository.xls"));
 
            //Create Workbook instance holding reference to .xlsx file
            HSSFWorkbook workbook = new HSSFWorkbook(file);
 
            //Get first/desired sheet from the workbook
            HSSFSheet sheet = workbook.getSheetAt(0);
 
            //Iterate through each rows one by one
            Iterator<Row> rowIterator = sheet.iterator();
            while (rowIterator.hasNext()) 
            {
                Row row = rowIterator.next();
               
                if(row.getRowNum()==0)
                		continue;
                //For each row, iterate through all the columns
                Iterator<Cell> cellIterator = row.cellIterator();
                 
                while (cellIterator.hasNext()) 
                {
                    Cell cell = cellIterator.next();
                                 
                            System.out.print(cell.getStringCellValue() + "\t\t\t");
                 
                }
                System.out.println("");
            }
            file.close();
            
            // write the content into xml file
	         TransformerFactory transformerFactory = TransformerFactory.newInstance();
	         Transformer transformer = transformerFactory.newTransformer();
	         DOMSource source = new DOMSource(doc);
	         StreamResult result = new StreamResult(new File("C:\\Users\\mamta_alane\\Documents\\cars1.xml"));
	         transformer.transform(source, result);
        } 
        catch (Exception e) 
        {
            e.printStackTrace();
        }
		
		CreateXMLFile obj=new CreateXMLFile();
		obj.createXML();
		
	}

}
