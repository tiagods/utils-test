package com.prolink.utilitarios.excel.contabil;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.*;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class Carteira {
	private StringBuilder builder = new StringBuilder();
    private String excelFilePath = "D://Carteirta Clientes - Bia 2018.xlsx";
    
	public static void main(String[] args) {
		new Carteira().iniciar();
	}
	private void iniciar() {
		List<Cliente> linhas = lerPlanilha();
		executar(linhas);
	}
	
	private List<Cliente> lerPlanilha() {
		try {
	    	List<Cliente> lista = new ArrayList<>();
	    	FileInputStream inputStream = new FileInputStream(new File(excelFilePath));
	        XSSFWorkbook workbook = (XSSFWorkbook) WorkbookFactory.create(inputStream);
	        XSSFSheet sheet = workbook.getSheetAt(1);
	
	        XSSFSheet sheetFinal = workbook.createSheet();
	        
	        int lastLine = 0;
	        Iterator<Row> rows = sheet.iterator();
	        while(rows.hasNext()) {
	        	Row row = rows.next();
	        	Row plan2 = sheetFinal.createRow(lastLine);
	        	plan2.createCell(0).setCellValue(row.getCell(0).getNumericCellValue());
	        	plan2.createCell(1).setCellValue(row.getCell(1).getStringCellValue());
	        	plan2.createCell(2).setCellValue(row.getCell(2).getStringCellValue());
	        	plan2.createCell(3).setCellValue(row.getCell(3).getStringCellValue());
	        	plan2.createCell(4).setCellValue(row.getCell(4).getStringCellValue());
	        	plan2.createCell(5).setCellValue(row.getCell(5).getStringCellValue());
//	        	for(int i = 5; i<17;i++) {
//	        		XSSFCell cell = (XSSFCell) plan2.createCell(i,CellType.FORMULA);
//	        		//cell.setCellValue("1");
//	        		cell.setCellFormula("SUM("+letra(i)+(plan2.getRowNum()+2)+":"
//	        		+letra(i)+(plan2.getRowNum()+6)+")/5");
//	        		CellStyle style = workbook.createCellStyle();
//	        		style.setDataFormat(workbook.createDataFormat().getFormat("0%"));
//	        		cell.setCellStyle(style);
//	        	}
	        	Row rowFolha = sheetFinal.createRow(plan2.getRowNum()+1);
	        	rowFolha.createCell(5).setCellValue("FOLHA");
	        	
	        	Row rowFiscal = sheetFinal.createRow(plan2.getRowNum()+2);
	        	rowFiscal.createCell(5).setCellValue("FISCAL");
	        	
	        	Row rowMov = sheetFinal.createRow(plan2.getRowNum()+3);
	        	rowMov.createCell(5).setCellValue("MOVIMENTO FINANCEIRO");
	        	
	        	Row rowClas = sheetFinal.createRow(plan2.getRowNum()+4);
	        	rowClas.createCell(5).setCellValue("CLASSIFICAÇÃO/DIGITAÇÃO");
	        	
	        	Row rowConc = sheetFinal.createRow(plan2.getRowNum()+5);
	        	rowConc.createCell(5).setCellValue("CONCILIADO");
	        	
	        	lastLine = rowConc.getRowNum()+1;
	        }
	        FileOutputStream outFile =new FileOutputStream(new File("d://Output.xlsx"));
            workbook.write(outFile);
	        workbook.close();
	        return lista;
    	}catch (IOException | EncryptedDocumentException
                | InvalidFormatException ex) {
            ex.printStackTrace();
            builder.append("#Falha ao criar a planilha do cliente ");
            return null;
        }
	}
	private String letra(int letra) {
		switch(letra) {
			case 5: return "F";
			case 6: return "G";
			case 7: return "H";
			case 8: return "I";
			case 9: return "J";
			case 10: return "K";
			case 11: return "L";
			case 12: return "M";
			case 13: return "N";
			case 14: return "O";
			case 15: return "P";
			case 16: return "Q";
			default: return null;
		}
	}
	private void executar(List<Cliente> clientes) {
		XSSFWorkbook wb = new XSSFWorkbook();
	    FileOutputStream fileOut = null;
		try {
			fileOut = new FileOutputStream("workbook.xlsx");
			XSSFSheet sheet = wb.createSheet("Relação");
//			sheet.create
//			wb.write(fileOut);
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		} finally {
		    try {
				fileOut.close();
			} catch (IOException e) {
				e.printStackTrace();
			}
		}
	}
	public class Cliente{
		String codigo;
		String perfil;
		String sistema;
		String tributacao;
		String nome;
		/**
		 * @return the codigo
		 */
		public String getCodigo() {
			return codigo;
		}
		/**
		 * @param codigo the codigo to set
		 */
		public void setCodigo(String codigo) {
			this.codigo = codigo;
		}
		/**
		 * @return the perfil
		 */
		public String getPerfil() {
			return perfil;
		}
		/**
		 * @param perfil the perfil to set
		 */
		public void setPerfil(String perfil) {
			this.perfil = perfil;
		}
		/**
		 * @return the sistema
		 */
		public String getSistema() {
			return sistema;
		}
		/**
		 * @param sistema the sistema to set
		 */
		public void setSistema(String sistema) {
			this.sistema = sistema;
		}
		/**
		 * @return the tributacao
		 */
		public String getTributacao() {
			return tributacao;
		}
		/**
		 * @param tributacao the tributacao to set
		 */
		public void setTributacao(String tributacao) {
			this.tributacao = tributacao;
		}
		/**
		 * @return the nome
		 */
		public String getNome() {
			return nome;
		}
		/**
		 * @param nome the nome to set
		 */
		public void setNome(String nome) {
			this.nome = nome;
		}	
	}
	
}
