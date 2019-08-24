package com.prolink.utilitarios.excel.fiscal;

import java.io.File;

import org.apache.commons.mail.DefaultAuthenticator;
import org.apache.commons.mail.EmailAttachment;
import org.apache.commons.mail.HtmlEmail;

public class SendMail {
	public boolean enviaAlerta(String message, String[] conta, String titulo, File arquivo, String fileName){
	    HtmlEmail email = new HtmlEmail();
	    email.setHostName("email-ssl.com.br");
	    email.setSmtpPort(587);

	    email.setAuthenticator( new DefaultAuthenticator( "comunicadoprolink@prolinkcontabil.com.br" ,  "plkc2004" ) );
	    
	    try {
	        email.setFrom( "comunicadoprolink@prolinkcontabil.com.br" , "Comunicado \\ Prolink Contabil");
	        email.setDebug(false);
	        
	        email.setSubject( titulo );

	        EmailAttachment anexo = new EmailAttachment();
	        anexo.setPath(arquivo.getAbsolutePath());//pegando o anexo
	        anexo.setDisposition(EmailAttachment.ATTACHMENT);
	        anexo.setName(fileName);//renomeando o arquivo pdf

	        email.attach(anexo);
	        
	        email.setHtmlMsg(message);
	        for(String c : conta) email.addTo(c);
	        
	        email.addBcc("robison.chan.tong@prolinkcontabil.com.br");
	        email.addBcc("tiago.dias@prolinkcontabil.com.br");
	        email.send();
	        return true;
	    } catch (Exception e) {
	    	e.printStackTrace();
	        return false;
	    } 
	    }
}
