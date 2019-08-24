package com.prolink.utilitarios.excelsimplesnacional;

import java.io.*;

import net.sourceforge.tess4j.ITesseract;
import net.sourceforge.tess4j.Tesseract;
import net.sourceforge.tess4j.TesseractException;

public class ExtractPDF {

	public static void main(String[] args) {
		// TODO Auto-generated method stub
		 
        File imageFile = new File("0510_REC.pdf");
        ITesseract instance = new Tesseract();
        instance.setLanguage("por");
        try {
            String result = instance.doOCR(imageFile);
            System.out.println(result);
            
            try {
            	FileWriter file = new FileWriter(new File("result.txt"));
            	file.write(result);
            	file.flush();
            	file.close();
            }catch (Exception e) {
			}
        } catch (TesseractException e) {
            System.out.println(e.getMessage());
        }
	}

}
