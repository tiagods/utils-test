package com.prolink.utilitarios.parametros.planocontasfiscal;

import java.io.*;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.nio.file.StandardCopyOption;
import java.util.ArrayList;
import java.util.List;

import com.prolink.utilitarios.distribuicao.DistribuicaoModel;

import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;

public class CopiaDoPlano {
	static StringBuilder builder = new StringBuilder();
	private File original = new File(System.getProperty("user.dir")+"/CContI.btr");
	
	public static void main(String[] args) {
		new CopiaDoPlano();
	}
	public CopiaDoPlano() {
		try {
			if(original.exists()) {
				List<DistribuicaoModel> lista = getList(System.getProperty("user.dir")+"/clientes-migracao-lista.xls");
				lista.forEach(c->{
					File path = new File("\\\\plkserver\\Sistemas\\Phoenix\\Empresas");
					
					boolean existeDois = false;
					boolean existeUm = false;
					String codigo_cliente = c.getCod_cli();
					if(codigo_cliente.length()<4){
						String zero = "";
						for(int v = 0; v<4-codigo_cliente.length(); v++){
							zero+="0";
						}
						codigo_cliente = zero+codigo_cliente;
					}
					
					File fileInicio = null;
					File fileRename = null;
					existeUm = new File(path+"\\"+c.getCod_cli()).exists();
					existeDois = c.getCod_cli().length()==4?false:new File(path+"\\"+codigo_cliente).exists();
					
					if(existeUm && existeDois) {
						builder.append(c.getCod_cli()+" - Existe dois clientes com o mesmo id");
					}
					else if(existeUm) {
						//renomear
						fileInicio = new File("\\\\plkserver\\Sistemas\\Phoenix\\Empresas\\"+c.getCod_cli()+"\\0\\CContI.btr");
						fileRename = new File("\\\\plkserver\\Sistemas\\Phoenix\\Empresas\\"+c.getCod_cli()+"\\0\\CContIOld.btr");
						
						System.out.println(codigo_cliente+"-"+(original.lastModified()==fileInicio.lastModified()));
						
						if(!fileInicio.exists())
							copiar(original,fileInicio,builder);
						else if(original.lastModified()!=fileInicio.lastModified()) {
							if(copiar(fileInicio,fileRename,builder)) {
								copiar(original,fileInicio,builder);
							}
						}
						else {
							builder.append(c.getCod_cli()+"-Arquivo já existe");
							builder.append(System.getProperty("line.separator"));
						}
						//copia oficial
					}
					else if(existeDois) {
						//renomear

						fileInicio = new File("\\\\plkserver\\Sistemas\\Phoenix\\Empresas\\"+codigo_cliente+"\\0\\CContI.btr");
						fileRename = new File("\\\\plkserver\\Sistemas\\Phoenix\\Empresas\\"+codigo_cliente+"\\0\\CContIOld.btr");
						System.out.println(codigo_cliente+"-"+(original.lastModified()==fileInicio.lastModified()));

						if(!fileInicio.exists())
							copiar(original,fileInicio,builder);
						else if(original.lastModified()!=fileInicio.lastModified()) {
							if(copiar(fileInicio,fileRename,builder)) {
								copiar(original,fileInicio,builder);
							}
						}
						else {
							builder.append(c.getCod_cli()+"-Arquivo já existe");
							builder.append(System.getProperty("line.separator"));
						}
//						if(copiar(fileInicio,fileRename,builder)) {
//							copiar(original,fileInicio,builder);
//						}
						//copia oficial
					}
				});
			}
			else {
				System.out.println("Falta Arquivo CContI.btr");
			}
			System.out.println("Processo concluido com sucesso");
			gerarArquivo();
		}catch(IOException e) {
			e.printStackTrace();
		} catch (BiffException ex) {
			ex.printStackTrace();
		}
	}
	private boolean copiar(File de, File para, StringBuilder builder){
		Path pathI = Paths.get(de.getAbsolutePath());
        Path pathO = Paths.get(para.getAbsolutePath());
        try{
        	Files.copy(pathI, pathO, StandardCopyOption.REPLACE_EXISTING);
        	builder.append("Arquivo copiado com sucesso! "+de.getAbsolutePath()+"-"+para.getAbsolutePath());
			builder.append(System.getProperty("line.separator"));
        	System.out.println("Arquivo copiado com sucesso! "+de.getAbsolutePath()+"-"+para.getAbsolutePath());
			return true;
        }catch (IOException e) {
			e.printStackTrace();
			builder.append("Erro ao copiar o arquivo "+de.getName()+"="+e.getMessage()+"-"+para.getAbsolutePath());
			builder.append(System.getProperty("line.separator"));
			System.out.println("Erro ao copiar o arquivo "+de.getName()+"="+e.getMessage()+"-"+para.getAbsolutePath());
			return false;
		}
	}
	private void gerarArquivo() throws IOException{
		FileWriter fw = new FileWriter(new File("resultado.txt"));
		fw.write(builder.toString());
		fw.close();
		fw.close();
	}
	public List<DistribuicaoModel> getList(String caminho) throws IOException, BiffException{
		Workbook workbook = null;
		try{
			FileInputStream file = new FileInputStream(new File(caminho));
			workbook = Workbook.getWorkbook(file);
			List<DistribuicaoModel> lista = new ArrayList<>();
			Sheet sheet = workbook.getSheet(0);
			int numberRows = sheet.getRows();
			System.out.println(numberRows);
			int i = 1;
			do{
				DistribuicaoModel dist = new DistribuicaoModel();
				dist.setCod_cli(getValue(sheet, i, 0));
				lista.add(dist);
				i++;
			}while(i<numberRows);
			workbook.close();
			return lista;
		}catch(IOException e){
			e.printStackTrace();
			return null;
		}
	}
	private String getValue(Sheet sheet, int line, int column){
		Cell cell = sheet.getCell(column,line);
		if(cell==null) return "";
		return sheet.getCell(column,line).getContents().trim();
	}
}
