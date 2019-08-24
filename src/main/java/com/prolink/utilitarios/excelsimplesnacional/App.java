package com.prolink.utilitarios.excelsimplesnacional;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.util.Comparator;
import java.util.HashSet;
import java.util.Iterator;
import java.util.Set;
import java.util.TreeSet;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFFormulaEvaluator;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import com.documents4j.api.DocumentType;
import com.documents4j.api.IConverter;
import com.documents4j.job.LocalConverter;
 
/**
 * This program illustrates how to update an existing Microsoft Excel document.
 * Append new rows to an existing sheet.
 *
 * @author www.codejava.net
 *
 */
public class App {
 
    public static void main(String[] args) {
    	new App().iniciar();
    }
    private StringBuilder builder = new StringBuilder();
    private String excelFilePath = "D://SN-COMPARATIVO_SIMPLIFICADO (05-01-2018)).xls";
    
    public void iniciar() {
    	Set<Integer> lista = listarClientes();
    	lista.forEach(id->{
    		int tamanho = String.valueOf(id).length();
    		String newId = tamanho == 4? String.valueOf(id):
    			(tamanho==3?"0"+id:
    				(tamanho==2?"00"+id:"000"+id));
    		File arquivo  = preencherNovaPlanilha(id,new File("I:\\COMERCIAL-SN 2018\\tabelas\\"+newId+" - Simples Nacional.xls"));
    		if(arquivo!=null) builder.append(arquivo.getName()).append(" gerado com sucesso!");
    		builder.append(System.getProperty("line.separator"));
    	});
    	lista.forEach(id->{
    		int tamanho = String.valueOf(id).length();
    		String newId = tamanho == 4? String.valueOf(id):
    			(tamanho==3?"0"+id:
    				(tamanho==2?"00"+id:"000"+id));
    		File file = new File("I:\\COMERCIAL-SN 2018\\tabelas\\"+newId+" - Simples Nacional.xls");
    		if(file.exists()) {
    			Conversor converter = new Conversor();
    			converter.convertToPdf(builder, file, new File("I:\\COMERCIAL-SN 2018\\tabelas\\"+newId+" - Simples Nacional.pdf"));
    		}
    		builder.append(System.getProperty("line.separator"));

    	});
    	try {
    		FileWriter writer = new FileWriter(new File("I:\\COMERCIAL-SN 2018\\tabelas\\resultado.txt"));
    		writer.write(builder.toString());
    		writer.flush();
    		writer.close();
    	}catch(Exception e) {
    		e.printStackTrace();
    	}
    }
    public Set<Integer> listarClientes(){
    	try {
	    	Set<Integer> lista = new TreeSet<>();
	    	FileInputStream inputStream = new FileInputStream(new File(excelFilePath));
	        HSSFWorkbook workbook = (HSSFWorkbook) WorkbookFactory.create(inputStream);
	
	        HSSFSheet sheet = workbook.getSheet("Clientes");
	        Iterator<Row> rows = sheet.iterator();
	        while(rows.hasNext()) {
	        	Row row = rows.next();
	        	if(row.getRowNum()==0) continue;
	        	Cell cellCodigo = row.getCell(0);
	        	Cell cellFaturamento = row.getCell(8);
	        	try {
	        		double value = cellCodigo.getNumericCellValue();
	        		double faturamento= cellFaturamento.getNumericCellValue();
	        		if(faturamento==0) builder.append(value).append("= Faturamento zerado\n");
	        		else lista.add((int)value);
	        	}catch(Exception e) {
	        		e.printStackTrace();
	        	}
	        }
	        return lista;
    	}catch (IOException | EncryptedDocumentException
                | InvalidFormatException ex) {
            ex.printStackTrace();
            builder.append("#Falha ao criar a planilha do cliente ");
            return null;
        }
    }
    public File preencherNovaPlanilha(int clienteId,File arquivo) {
    	 try {
            FileInputStream inputStream = new FileInputStream(new File(excelFilePath));
            HSSFWorkbook workbook = (HSSFWorkbook) WorkbookFactory.create(inputStream);
 
            HSSFSheet sheet = workbook.getSheetAt(0);
            
            HSSFCell id = sheet.getRow(2).getCell(2);
            id.setCellValue(clienteId);
            
            //HSSFCell nome = sheet.getRow(2).getCell(4);
            //nome.setCellValue("Empresa Teste Ltda");
            
            //HSSFCell faturamento = sheet.getRow(4).getCell(3);
            //faturamento.setCellValue(100000);
            //faturamento.setCellType(CellType.NUMERIC);
            
            //HSSFCell prolabore = sheet.getRow(8).getCell(3);
            //prolabore.setCellValue(1200);
            //prolabore.setCellType(CellType.NUMERIC);
            
            HSSFFormulaEvaluator.evaluateAllFormulaCells(workbook);
            
            Double value = processar(sheet,clienteId);
            
            HSSFCell minimo = sheet.getRow(16).getCell(8);
            minimo.setCellType(CellType.NUMERIC);
            minimo.setCellValue(value);
            
            HSSFCell minimo2 = sheet.getRow(18).getCell(8);
            minimo2.setCellType(CellType.NUMERIC);
            minimo2.setCellValue(value);
            
            HSSFFormulaEvaluator.evaluateAllFormulaCells(workbook);
            inputStream.close();
            
            FileOutputStream outputStream = new FileOutputStream(arquivo);
            workbook.write(outputStream);
            workbook.close();
            outputStream.close();
            
            return arquivo;
        } catch (IOException | EncryptedDocumentException
                | InvalidFormatException ex) {
            ex.printStackTrace();
            builder.append("#Falha ao criar a planilha do cliente ").append(clienteId);
            return null;
        }
        
    }
    public Double processar(HSSFSheet sheet, int clienteId) {
    	double result = 0;
    	int i = 1;
    	builder.append(clienteId);
    	builder.append("=\n");
    	while(i<=6) {
    		Double d = sheet.getRow(16).getCell(i).getNumericCellValue();
    		if(d>0) { 
    			if(result==0 || result>d) result = d;
    		}
    		builder.append(d).append("\t");
    		i++;
    	}
    	builder.append("-> O menor é ").append(result);
		return result;
    }    
 
}