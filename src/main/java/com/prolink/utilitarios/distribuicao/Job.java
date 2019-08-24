package com.prolink.utilitarios.distribuicao;

import java.io.File;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.List;
import java.util.stream.Collectors;

public class Job {
	public static void main(String[] args) {
		String caminhoPlanilha = System.getProperty("user.dir")+"/Planilha_distribuicao 2018_2017.xls";
		String modelo = System.getProperty("user.dir")+"/ModeloDistribuicao.doc";
		ProcessarPlanilha processar = new ProcessarPlanilha();
		List<DistribuicaoModel> listaDeSocios = null;
		try {
			listaDeSocios = processar.getList(caminhoPlanilha);
			System.out.println(listaDeSocios.size());
			
			List<DistribuicaoModel> novaLista = listaDeSocios.stream().filter(c->
			c.getCod_cli().equals("371")

					).collect(Collectors.toList());
			for(int i = 0 ; i<novaLista.size(); i++){
				DistribuicaoModel dist = novaLista.get(i);
	
				String[] socioName = dist.getNome_do_socio().split(" ");
				String codigo_cliente = dist.getCod_cli();
				if(codigo_cliente.length()<4){
					String zero = "";
					for(int v = 0; v<4-codigo_cliente.length(); v++){
						zero+="0";
					}
					codigo_cliente = zero+codigo_cliente;
				}
	
				String temp = new SimpleDateFormat("HHmm").format(new Date());
	
				String file = new File("c:/temp/"+codigo_cliente+" - CARTA DE DISTRIBUICAO "+socioName[0]+".doc").exists()?
						"c:/temp/"+codigo_cliente+" - CARTA DE DISTRIBUICAO "+socioName[0]+" "+temp+".doc":"c:/temp/"+codigo_cliente+" - CARTA DE DISTRIBUICAO "+socioName[0]+".doc";
				EditorDoc doc = new EditorDoc(modelo, file);
				doc.replaceTag("{cpf}", dist.getCpf());
				doc.replaceTag("{razao_social}", dist.getRazao_social());
				doc.replaceTag("{cnpj}", dist.getCnpj());
				doc.replaceTag("{status}", dist.getStatus());
				doc.replaceTag("{cod_cli}", dist.getCod_cli());
				doc.replaceTag("{nome_do_socio}", dist.getNome_do_socio());
				doc.replaceTag("{pro-labore}", dist.getPro_labore());
				doc.replaceTag("{inss_11}", dist.getInss_11());
				doc.replaceTag("{irrf}", dist.getIrrf().trim().equals("")?"0,00":dist.getIrrf());
				doc.replaceTag("{tt_distribuicao}", dist.getTt_distribuicao());
				doc.replaceTag("{percentual}", dist.getPercentual());
				doc.replaceTag("{valor_capital}", dist.getValor_capital());
				doc.replaceTag("{dt_admissao}", dist.getDt_admissao());
				doc.replaceTag("{contab_distrib}", dist.getContab_distrib());
				doc.write();
				//if(i==5) break;
//				Conversor converter = new Conversor();
//				converter.convertToPdf(new StringBuilder(), new File(file), new File(file.replace(".doc", ".pdf")));
			}
		} catch (Exception e) {
			e.printStackTrace();
		}
		System.exit(0);
	}
}
