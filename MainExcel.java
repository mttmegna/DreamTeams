package it.jac.corsojava.io;

import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class MainExcel {
	
	private static Logger log = LogManager.getLogger(MainExcel.class);

	public static void main(String[] args) {

		try (XSSFWorkbook workbook = new XSSFWorkbook()) {
		
			XSSFSheet sheet = workbook.createSheet("Datatypes in Java");
			Object[][] datatypes = { 
					{ "Datatype", "Type", "Size(in bytes)" }, 
					{ "int", "Primitive", 2 },
					{ "float", "Primitive", 4 }, 
					{ "double", "Primitive", 8 }, 
					{ "char", "Primitive", 1 },
					{ "String", "Non-Primitive", "No fixed size" } };
	
			int rowNum = 0;
			log.info("Creating excel");
	
			for (Object[] datatype : datatypes) {
				
				Row row = sheet.createRow(rowNum++);
				int colNum = 0;
				
				for (Object field : datatype) {
					
					Cell cell = row.createCell(colNum++);
					
					if (field instanceof String) {
						cell.setCellValue((String) field);
					} else if (field instanceof Integer) {
						cell.setCellValue((Integer) field);
					}
				}
			}

			FileOutputStream outputStream = new FileOutputStream("c:\\temp\\work\\excelExample.xlsx");
			workbook.write(outputStream);
			
		} catch (IOException e) {
			
			log.error("Error creating excel file", e);
		}

		log.info("Done");
	}
}
