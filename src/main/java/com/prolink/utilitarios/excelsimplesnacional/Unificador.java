package com.prolink.utilitarios.excelsimplesnacional;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileWriter;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Iterator;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class Unificador {
	
	static StringBuilder builder = new StringBuilder();
	
	int num = 2;
	
	public static void main(String[] args) {
		try {
			FileInputStream inputStream = new FileInputStream(new File("C:\\Users\\Tiago\\Desktop\\Fornecedores.xls"));
			HSSFWorkbook workbook = (HSSFWorkbook) WorkbookFactory.create(inputStream);
			Iterator<Sheet> sheets = workbook.sheetIterator();
			while(sheets.hasNext()) {
				Sheet sheet = sheets.next();
				lerPlanilha(sheet,builder);
			}
		} catch (EncryptedDocumentException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (InvalidFormatException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		try {
			
    		FileWriter writer = new FileWriter(new File("d:\\resultado.csv"));
    		writer.write(builder.toString());
    		writer.flush();
    		writer.close();
    	}catch(Exception e) {
    		e.printStackTrace();
    	}
	}
	public static void lerPlanilha(Sheet sheet, StringBuilder builder) {
		try {
			Iterator<Row> rows = sheet.iterator();
			while (rows.hasNext()) {
				Row row = rows.next();
				Iterator<Cell> cells = row.cellIterator();
				builder.append(sheet.getSheetName());
				builder.append(";");
				while(cells.hasNext()) {
					Cell cell = cells.next();
					String valor = readingCell((HSSFCell)cell);
					if(row.getRowNum()>1 && !valor.equals("") || sheet.getSheetName().equalsIgnoreCase("table 2")) {
						builder.append(valor);
						builder.append(";");
					}
				}
				builder.append("\n");
			}
		} catch (Exception ex) {
			ex.printStackTrace();
		}
	}
	private static String readingCell(HSSFCell celula){ //metodo usado para tratar as celulas
		if(celula==null)
			return "";
		switch (celula.getCellType()) {
			case HSSFCell.CELL_TYPE_NUMERIC:
				if(DateUtil.isCellDateFormatted(celula))
					return new SimpleDateFormat("dd/MM/yyyy").format(celula.getDateCellValue());//campo do tipo data formatando no ato
				else{
					return String.valueOf(celula.getNumericCellValue());//campo do tipo numerico, convertendo double para long
				}
			case HSSFCell.CELL_TYPE_STRING:
				return String.valueOf(celula.getStringCellValue());//conteudo do tipo texto
			case HSSFCell.CELL_TYPE_BOOLEAN:    
				return "";
			case HSSFCell.CELL_TYPE_BLANK:
				return "";
			case HSSFCell.CELL_TYPE_FORMULA:
				return "";
			default:
				return "";
		}
    }
	
}
