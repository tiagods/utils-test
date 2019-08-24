package com.prolink.utilitarios.excel.fiscal;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Message {
	StringBuilder builder = new StringBuilder();
	
	StringBuilder log = new StringBuilder();
	int inicial = 0;
	
	private String excelFilePath = "D://Estoque2017.xlsx";
	
	public static void main(String[] args) {
		new Message().start();
	}
	private void start() {
		getMessage();
		List<Cliente> lista = listarClientes();
		lista.forEach(c->{
				int tamanho = String.valueOf(c.getId()).length();
	    		String newId = tamanho == 4? String.valueOf(c.getId()):
	    			(tamanho==3?"0"+c.getId():
	    				(tamanho==2?"00"+c.getId():"000"+c.getId()));
	    		
	    		SendMail send = new SendMail();
	    		String[] conta = c.getEmail().split(";");
	    		//String[] conta = new String[] {"suporte.ti@prolinkcontabil.com.br"};
	    		if(inicial<90){
	    			boolean sucess = send.enviaAlerta(builder.toString(), conta, newId+"-Carta de solicitação de estoque ref. 2017", new File("D://inventario.xls"), newId+"-Estoque 2017.xls");
	    			inicial+=conta.length+2;
	    			System.out.println("Enviado comunicado "+newId+"-"+sucess+"-"+inicial);
		    		log.append("Enviado comunicado "+newId+"="+sucess);
	    		}
		});
		try {
    		FileWriter writer = new FileWriter(new File("D:\\resultado.txt"));
    		writer.write(log.toString());
    		writer.flush();
    		writer.close();
    	}catch(Exception e) {
    		e.printStackTrace();
    	}
	}
	 public List<Cliente> listarClientes(){
	    	try {
	    		List<Cliente> lista = new ArrayList<>();
		    	FileInputStream inputStream = new FileInputStream(new File(excelFilePath));
		        XSSFWorkbook workbook = (XSSFWorkbook) WorkbookFactory.create(inputStream);
		        XSSFSheet sheet = workbook.getSheetAt(0);
		        Iterator<Row> rows = sheet.iterator();
		        while(rows.hasNext()) {
		        	Row row = rows.next();
		        	if(row.getRowNum()==0) continue;
		        	Cliente cliente = new Cliente();
		        	cliente.setId((int)row.getCell(0).getNumericCellValue());
		        	cliente.setEmpresaNome(row.getCell(4).getStringCellValue());
		        	cliente.setNome(row.getCell(5).getStringCellValue());
		        	cliente.setEmail(row.getCell(7).getStringCellValue());
		        	lista.add(cliente);
		        }
		        return lista;
	    	}catch (IOException | EncryptedDocumentException
	                | InvalidFormatException ex) {
	            ex.printStackTrace();
	            builder.append("#Falha ao criar a planilha do cliente ");
	            return null;
	        }
	    }
	
	private void getMessage() {
		builder.append("<html>");
		builder.append("	<body>");
		builder.append("		<div style=\"text-align: left;\">");
		builder.append("			<p>");
		builder.append("				<strong>Prezado Cliente,</strong></p>");
		builder.append("			<p>");
		builder.append("				&nbsp;</p>");
		builder.append("			<p>");
		builder.append("				Com o encerramento do exerc&iacute;cio a contabilidade precisa das informa&ccedil;&otilde;es de estoques com data de 31/12/2017.</p>");
		builder.append("			<p>");
		builder.append("				Esta informa&ccedil;&atilde;o &eacute; importante para o fechamento de balan&ccedil;o e est&aacute; sendo exigido pela receita federal e secretaria da fazenda para que os dados sejam informados, conforme o caso:</p>");
		builder.append("			<p>");
		builder.append("				&nbsp;</p>");
		builder.append("			<p>");
		builder.append("				<strong><u>Empresas do Simples Nacional </u></strong></p>");
		builder.append("			<p>");
		builder.append("				&nbsp;</p>");
		builder.append("			<p>");
		builder.append("				&bull;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; Informa&ccedil;&otilde;es pertinentes &agrave; declara&ccedil;&atilde;o anual DEFIS referente 2017 e apresenta&ccedil;&atilde;o em caso de fiscaliza&ccedil;&atilde;o do livro registro de invent&aacute;rio.</p>");
		builder.append("			<p>");
		builder.append("				&nbsp;</p>");
		builder.append("			<p>");
		builder.append("				<strong><u>Empresas que apuram ICMS sobre regime normal (RPA) </u></strong></p>");
		builder.append("			<p>");
		builder.append("				&nbsp;</p>");
		builder.append("			<p>");
		builder.append("				&bull;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; Informa&ccedil;&otilde;es pertinentes ao Sped Fiscal referente 02/2018, com dados do exerc&iacute;cio anterior, devidamente detalhados em arquivo magn&eacute;tico.</p>");
		builder.append("			<p>");
		builder.append("				&bull;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; Apresenta&ccedil;&atilde;o de informa&ccedil;&otilde;es na ECF-Escritura&ccedil;&atilde;o Cont&aacute;bil Fiscal.</p>");
		builder.append("			<p>");
		builder.append("				&nbsp;</p>");
		builder.append("			<p>");
		builder.append("				Assim solicita-se a gentileza de nos enviar a sua rela&ccedil;&atilde;o de estoques conforme o layout planilhado em Excel anexo <span style=\"color:#ff0000;\"><u><strong>at&eacute; o dia 28/02/2018.</strong></u></span></p>");
		builder.append("			<p>");
		builder.append("				&nbsp;</p>");
		builder.append("			<p>");
		builder.append("				<strong>Importante: caso n&atilde;o haja&nbsp; mercadorias em estoque, nos informar.</strong></p>");
		builder.append("			<p>");
		builder.append("				&nbsp;</p>");
		builder.append("			<p>");
		builder.append("				Reiteramos a necessidade destas informa&ccedil;&otilde;es, pois como &eacute; de conhecimento a intelig&ecirc;ncia atual do Fisco em rela&ccedil;&atilde;o a cruzamentos de informa&ccedil;&otilde;es est&aacute; cada dia mais preciso e din&acirc;mico sendo que, na omiss&atilde;o ou dados incompletos nas obriga&ccedil;&otilde;es fiscais anuais est&atilde;o a gerar multas e outras penalidades.</p>");
		builder.append("			<p>");
		builder.append("				&nbsp;</p>");
		builder.append("			<p>");
		builder.append("				<strong>Favor encaminhar os arquivos para o endere&ccedil;o de e-mail: </strong></p>");
		builder.append("			<p>");
		builder.append("				&bull;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <a href=\"mailto:fiscal@prolinkcontabil.com.br\">fiscal@prolinkcontabil.com.br</a></p>");
		builder.append("			<p>");
		builder.append("				&bull;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; Assunto: ID-Estoque 2017</p>");
		builder.append("			<p>");
		builder.append("				Onde ID &eacute; seu c&oacute;digo de identifica&ccedil;&atilde;o em nossos controles. Caso tenha alguma d&uacute;vida, estamos &agrave; disposi&ccedil;&atilde;o para esclarecimentos.</p>");
		builder.append("			<p>");
		builder.append("				&nbsp;</p>");
		builder.append("			<p>");
		builder.append("				Atenciosamente,</p>");
		builder.append("			<p>");
		builder.append("				<strong>PLK-ProLink Assessoria Cont&aacute;bil Ltda-EPP</strong></p>");
		builder.append("		</div>");
		builder.append("		<p>");
		builder.append("			&nbsp;</p>");
		builder.append("	</body>");
		builder.append("</html>");
	}

	public class Cliente{
		private int id;
		private String nome;
		private String empresaNome;
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
		/**
		 * @return the empresaNome
		 */
		public String getEmpresaNome() {
			return empresaNome;
		}
		/**
		 * @param empresaNome the empresaNome to set
		 */
		public void setEmpresaNome(String empresaNome) {
			this.empresaNome = empresaNome;
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
