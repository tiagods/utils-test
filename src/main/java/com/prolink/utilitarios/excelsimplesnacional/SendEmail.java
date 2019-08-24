package com.prolink.utilitarios.excelsimplesnacional;

import java.io.File;

import org.apache.commons.mail.DefaultAuthenticator;
import org.apache.commons.mail.EmailAttachment;
import org.apache.commons.mail.EmailException;
import org.apache.commons.mail.HtmlEmail;

public class SendEmail {
	String message = "<html>\r\n" + 
			"	<head>\r\n" + 
			"		<title>Editor HTML Online</title>\r\n" + 
			"	</head>\r\n" + 
			"	<body>\r\n" + 
			"		<div style=\"text-align: left;\">\r\n" + 
			"			<p>\r\n" + 
			"				<em>Prezado(a) cliente, boa noite,</em></p>\r\n" + 
			"			<p>\r\n" + 
			"				&nbsp;</p>\r\n" + 
			"			<p>\r\n" + 
			"				<em>De acordo com as novas regras do Simples Nacional 2018, &eacute; poss&iacute;vel a uma empresa que hoje est&aacute; no regime tribut&aacute;rio de Lucro Presumido ser enquadrada no Simples Nacional.</em></p>\r\n" + 
			"			<p>\r\n" + 
			"				&nbsp;</p>\r\n" + 
			"			<p>\r\n" + 
			"				<em>A sua empresa, atualmente, est&aacute; inserida no regime tribut&aacute;rio do Lucro Presumido. A Prolink realizou uma an&aacute;lise para lhe informar se &eacute; poss&iacute;vel migrar para o Simples Nacional e se o mesmo &eacute; conveniente para sua empresa.</em></p>\r\n" + 
			"			<p>\r\n" + 
			"				&nbsp;</p>\r\n" + 
			"			<p>\r\n" + 
			"				<em>Com este prop&oacute;sito, segue anexo um explicativo baseado em sua m&eacute;dia de faturamento, demonstrando se &eacute; vi&aacute;vel ou n&atilde;o esta mudan&ccedil;a, levando em considera&ccedil;&atilde;o o seu ramo de atividade.</em></p>\r\n" + 
			"			<p>\r\n" + 
			"				&nbsp;</p>\r\n" + 
			"			<p>\r\n" + 
			"				<em>Nota: uma vez no Simples Nacional, a avalia&ccedil;&atilde;o para recolhimento no Anexo III considera o faturamento m&eacute;dio dos &uacute;ltimos 12 meses e o pr&oacute;-labore do mesmo per&iacute;odo, deste modo, o enquadramento s&oacute; ir&aacute; acontecer ap&oacute;s o per&iacute;odo de um ano recolhendo pelo Anexo V, caso o pr&oacute;-labore represente 28% ou mais do faturamento da empresa. Lembrando ainda que, se o faturamento aumentar, ser&aacute; necess&aacute;rio ajustar o valor de seu pr&oacute;-labore para que a propor&ccedil;&atilde;o m&iacute;nima de 28% seja atingida.</em></p>\r\n" + 
			"			<p>\r\n" + 
			"				&nbsp;</p>\r\n" + 
			"			<p>\r\n" + 
			"				<em>Um de nossos colaboradores ir&aacute; lhe contatar em breve para sanar as d&uacute;vidas.</em></p>\r\n" + 
			"			<p>\r\n" + 
			"				<br />\r\n" + 
			"				<em>Cordialmente,</em></p>\r\n" + 
			"			<p>\r\n" + 
			"				<strong><em>Prolink Contabil</em></strong></p>\r\n" + 
			"		</div>\r\n" + 
			"		<p>\r\n" + 
			"			&nbsp;</p>\r\n" + 
			"	</body>\r\n" + 
			"</html>\r\n" + 
			"";
	
    public boolean enviaAlerta(String[] conta, String titulo, File arquivo){
    HtmlEmail email = new HtmlEmail();
    email.setHostName( "email-ssl.com.br" );
    email.setSmtpPort(587);

    email.setAuthenticator( new DefaultAuthenticator( "bip@prolinkcontabil.com.br" ,  "prolink10" ) );
    
    try {
        email.setFrom( "comunicadoprolink@prolinkcontabil.com.br" , "Comunicado Prolink \\ Prolink Contabil");
        email.setDebug(true);
        
        email.setSubject( titulo );

        EmailAttachment anexo = new EmailAttachment();
        anexo.setPath(arquivo.getAbsolutePath());//pegando o anexo
        anexo.setDisposition(EmailAttachment.ATTACHMENT);
        anexo.setName(arquivo.getName());//renomeando o arquivo pdf
//        
        email.attach(anexo);
        
        email.setHtmlMsg(message);
        for(String c : conta) email.addTo(c);
        email.addCc("victor.santos@prolinkcontabil.com.br");
        email.addBcc("tiago.dias@prolinkcontabil.com.br");
        email.send();
        return true;
    } catch (Exception e) {
    	e.printStackTrace();
        return false;
    } 
    }
    public static void main(String[] args) {
    	new SendEmail().enviaAlerta(new String[] {"suporte.ti@prolinkcontabil.com.br","alertas@prolinkcontabil.com.br"}, 
    			"2200 - Enquadramento Simples Nacional", new File("d://cartasandra.pdf"));
    }
}
