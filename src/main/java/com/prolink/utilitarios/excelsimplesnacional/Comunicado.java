package com.prolink.utilitarios.excelsimplesnacional;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.io.Serializable;
import java.util.HashSet;
import java.util.Iterator;
import java.util.Set;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class Comunicado {
	private StringBuilder builder = new StringBuilder();
    private String excelFilePath = "D://job.xls";
    
	public static void main(String[] args) {
		// TODO Auto-generated method stub
		new Comunicado().iniciar();
	}
	
	public void iniciar() {
		Set<Cliente> lista = listarClientes();
		if(!lista.isEmpty()) {
			lista.forEach(c->{
				Cliente cli = c;
				String status = "";
				Set<Integer> valores = new HashSet<>();
				valores.add(1835);
				valores.add(1938);
				valores.add(2015);
				valores.add(2212);
				valores.add(2268);
				valores.add(2275);
				valores.add(2308);
				valores.add(2319);

				boolean pular = !valores.contains(cli.getId());
				if(!pular) {
					if(cli.getEmail().trim().equals("*")) {
						status = "Sem e-mail cadastrado";
					}
					else {
						int tamanho = String.valueOf(cli.getId()).length();
			    		String newId = tamanho == 4? String.valueOf(cli.getId()):
			    			(tamanho==3?"0"+cli.getId():
			    				(tamanho==2?"00"+cli.getId():"000"+cli.getId()));
						File file = pegarArquivo(newId);
						if(file!=null) {
							SendEmail email = new SendEmail();
							String[] contas = cli.getEmail().split(";");
							//String[] contas = new String[] {"alertas@prolinkcontabil.com.br","suporte.ti@prolinkcontabil.com.br"};
							boolean sucess = email.enviaAlerta(contas, newId+" - Enquadramento Simples Nacional", file);
							if(sucess) status = "Enviado com sucesso";
							else status = "Não enviado";
						}
						else {
							status = "Arquivo não encontrado";
						}
					}
					try {
						FileWriter fw = new FileWriter("d:/bloco.csv",true);
						fw.write(cli.getId()+";"+cli.getEmail().replace(";", ",")+";"+status);
						fw.write(System.getProperty("line.separator"));
						fw.flush();
						fw.close();
					}catch(Exception e) {}
				}
			});
			
		}
	}
	public File pegarArquivo(String id) {
		File file = new File("D:\\tabelas\\"+id+" - Simples Nacional.pdf");
		if(file.exists()) return file;
		else return null;
	}
	
	public Set<Cliente> listarClientes(){
    	try {
	    	Set<Cliente> lista = new HashSet<>();
	    	FileInputStream inputStream = new FileInputStream(new File(excelFilePath));
	        HSSFWorkbook workbook = (HSSFWorkbook) WorkbookFactory.create(inputStream);
	
	        HSSFSheet sheet = workbook.getSheet("Clientes");
	        Iterator<Row> rows = sheet.iterator();
	        while(rows.hasNext()) {
	        	Cliente cliente = new Cliente();
	        	Row row = rows.next();
	        	Cell cellCodigo = row.getCell(0);
	        	Cell cellEmail = row.getCell(12);
	        	try {
	        		cliente.setId((int)cellCodigo.getNumericCellValue());
	        		cliente.setEmail(cellEmail.getStringCellValue()==null?"":cellEmail.getStringCellValue());
	        		if(cliente.getEmail().toString().length()==0) builder.append(cliente.getId()).append("= Email vazio\n");
	        		else lista.add(cliente);
	        	}catch(Exception e) {
	        		System.out.println("Erro - "+row.getRowNum());
	        		e.printStackTrace();
	        	}
	        }
	        return lista;
    	}catch (IOException | EncryptedDocumentException
                | InvalidFormatException ex) {
            ex.printStackTrace();
            return null;
        }
    }	
	public class Cliente implements Serializable{
		
		/**
		 * 
		 */
		private static final long serialVersionUID = 1L;
		private int id;
		private String email;
		/**
		 * @return the id
		 */
		public int getId() {
			return id;
		}
		/**
		 * @param id the id to set
		 */
		public void setId(int id) {
			this.id = id;
		}
		/**
		 * @return the email
		 */
		public String getEmail() {
			return email;
		}
		/**
		 * @param email the email to set
		 */
		public void setEmail(String email) {
			this.email = email;
		}
		
		
		
	}
}
