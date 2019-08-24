package com.prolink.utilitarios.excelsimplesnacional;

import java.io.File;

import org.artofsolving.jodconverter.OfficeDocumentConverter;
import org.artofsolving.jodconverter.office.DefaultOfficeManagerConfiguration;
import org.artofsolving.jodconverter.office.OfficeManager;

public class Conversor {
	public static void main(String[] args) {
		new Conversor().convertToPdf(new StringBuilder(), new File("I:\\COMERCIAL-SN 2018\\tabelas\\2309 - Simples Nacional.xls"), new File("I:\\COMERCIAL-SN 2018\\tabelas\\2309 - Simples Nacional.pdf"));
	}
	
	public void convertToPdf(StringBuilder builder, File local, File destino){
		// TODO Auto-generated method stub
		OfficeManager officeManager = null;
		try {
			officeManager = new DefaultOfficeManagerConfiguration()
					.setOfficeHome(new File("C:\\Program Files\\LibreOffice 5"))
					.buildOfficeManager();
			officeManager.start();
			
			OfficeDocumentConverter converter = new OfficeDocumentConverter(officeManager);
			converter.convert(local, destino);
			//converter.convert(new File("d:/SN-COMPARATIVO_SIMPLIFICADO (05-01-2018)).xls"), new File("d:/teste.pdf"));
			//converter.convert(new File("d:/0105 - CARTA DE DISTRIBUICAO SANDRA.doc"), new File("d:/cartasandra.pdf"));
			builder.append("Conversão PDF Concluida!");
		}catch (Exception e) {
			e.printStackTrace();
			builder.append("Erro na conversão pdf!");
		}
		finally {
			if(officeManager != null) {
				officeManager.stop();
			}
		}
	}
}
