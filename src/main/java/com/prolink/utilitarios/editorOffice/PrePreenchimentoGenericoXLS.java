package com.prolink.utilitarios.editorOffice;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;


public class PrePreenchimentoGenericoXLS {

	public static void main(String[] args) throws Exception{
		// TODO Auto-generated method stub
		String modelo = System.getProperty("user.dir")+"/Cronograma-de-Selecao.xls";
		try {
	        FileInputStream file = new FileInputStream(new File(modelo));
	        HSSFWorkbook workbook = new HSSFWorkbook(file);
			HSSFSheet sheet = workbook.getSheetAt(0);
			Iterator<Row> iteratorRow = sheet.rowIterator();
			while(iteratorRow.hasNext()) {
				Row row = iteratorRow.next();
				Iterator<Cell>iterator = row.cellIterator();
				while(iterator.hasNext()) {
					Cell cell = iterator.next();
					String as1 = cell.getStringCellValue();
					if (as1.contains("{cargo}")) {
						cell.setCellValue(as1.replace("{cargo}","Analista de Sistemas"));
					}
				}
			}
			file.close();
			FileOutputStream outFile = new FileOutputStream(
					new File(System.getProperty("user.dir") + "/" + "resultado.xls"));
			HSSFWorkbook workbook2 = new HSSFWorkbook(file);
			workbook2.write(outFile);
			outFile.close();
			System.out.println("Arquivo Excel editado com sucesso!");
	        
	    } catch (FileNotFoundException e) {
	        System.out.println("Arquivo Excel não encontrado!");
	    } catch (IOException e) {
	        System.out.println("Erro na edição do arquivo!");
	    }
	}

}
