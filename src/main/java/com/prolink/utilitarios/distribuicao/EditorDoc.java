package com.prolink.utilitarios.distribuicao;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.usermodel.Range;  

public class EditorDoc {
	private POIFSFileSystem fs = null;  
    private FileInputStream fis = null;  
    private HWPFDocument doc = null;  
      
    private String arquivoSaida;   // pasta de saida do arquivo   
    /** 
     * Cria o arquivo na pasta indicada 
     * @param arquivoEntrada - endereco do arquivo de entrada (modelo) 
     * @param arquivoSaida - endereco do arquivo de saida, onde vai ser salvo o relatório 
     */  
      
    public EditorDoc(String arquivoEntrada, String arquivoSaida){  
        this.arquivoSaida = arquivoSaida;  
          
        try {  
            fis = new FileInputStream(arquivoEntrada);  
            fs = new POIFSFileSystem(fis);  
            doc = new HWPFDocument(fs);    
        } catch (FileNotFoundException e) {  
            e.printStackTrace();  
        } catch (IOException e) {  
            e.printStackTrace();  
        }  
    }  
      
      
    /** 
     * Metodo responsavel por criar o arquivo na pasta informada como parametro na criacao do objeto 
     */  
      
    public void write(){  
          
        FileOutputStream fos;  
        try {  
            fos = new FileOutputStream(arquivoSaida);  
            doc.write(fos);  
            fis.close();  
            fos.close();  
            System.out.println("Arquivo Gerado com Sucesso!!");  
        } catch (FileNotFoundException e) {  
            // TODO Auto-generated catch block  
            e.printStackTrace();  
        } catch (IOException e) {  
            // TODO Auto-generated catch block  
            e.printStackTrace();  
        }  
    }  
      
  
    /** 
     * METODO USADO PARA SUBSTITUIR UMA TAG POR UMA STRING QUE O PROGRAMADOR DESEJAR 
     * @param tag 
     * @param novaPalavra 
     */  
    public void replaceTag(String tag, String novaPalavra){  
    	Range range = doc.getRange();  
    	range.replaceText(tag,novaPalavra);          
    }    
}
