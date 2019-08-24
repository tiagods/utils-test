package com.prolink.utilitarios.distribuicao;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;

public class ProcessarPlanilha {
	public List<DistribuicaoModel> getList(String caminho) throws Exception{
		Workbook workbook = null;
		try{
			FileInputStream file = new FileInputStream(caminho);
			workbook = Workbook.getWorkbook(file);
			List<DistribuicaoModel> lista = new ArrayList<>();
			Sheet sheet = workbook.getSheet(0);
			int numberRows = sheet.getRows();
			System.out.println(numberRows);
			int i = 1;
			do{
				DistribuicaoModel dist = new DistribuicaoModel();
				dist.setCod_cli(getValue(sheet, i, 0));
				dist.setApto(getValue(sheet, i, 1));
				dist.setDirf(getValue(sheet, i, 2));
				//dist.setResultado(getValue(sheet, i, 2));
				dist.setCarta(getValue(sheet, i, 3));
				dist.setAtiva_inat(getValue(sheet, i, 4));
				dist.setStatus(getValue(sheet, i, 5));
				dist.setRazao_social(getValue(sheet, i, 6));
				dist.setCnpj(getValue(sheet, i, 7));
				dist.setNome_do_socio(getValue(sheet, i, 8));
				dist.setPercentual(getValue(sheet, i, 9));
				dist.setVr_distribuido(getValue(sheet, i, 10));
				dist.setTt_distribuicao(getValue(sheet, i, 11));
				dist.setContab_distrib(getValue(sheet, i, 10));
				dist.setPro_labore(getValue(sheet, i, 12));
				dist.setInss_11(getValue(sheet, i, 13));
				dist.setIrrf(getValue(sheet, i, 14));
				dist.setCpf(getValue(sheet, i, 15));
				dist.setValor_capital(getValue(sheet, i, 16));
//				dist.setConstituicao(getValue(sheet, i, 19));
				dist.setDt_admissao(getValue(sheet, i, 24));
//				dist.setDtdemissao(getValue(sheet, i, 21));
//				dist.setFaturamento(getValue(sheet, i, 22));
//				dist.setSaldo_do_caixa(getValue(sheet, i, 23));
//				dist.setLucro_periodo(getValue(sheet, i, 24));
//				dist.setPercentual1(getValue(sheet, i, 25));
//				dist.setBase_distribuicao(getValue(sheet, i, 26));
//				dist.setDistribuicao(getValue(sheet, i, 27));
//				dist.setProlabore(getValue(sheet, i, 28));
//				dist.setIdsocio(getValue(sheet, i, 29));
//				dist.setCodigo(getValue(sheet, i, 30));
//				dist.setEndereco(getValue(sheet, i, 31));
//				dist.setNumero(getValue(sheet, i, 32));
//				dist.setComplemento(getValue(sheet, i, 33));
//				dist.setBairro(getValue(sheet, i, 34));
//				dist.setUf(getValue(sheet, i, 35));
//				dist.setCep(getValue(sheet, i, 36));
//				dist.setEmail(getValue(sheet, i, 37));
//				dist.setQtde_filhos(getValue(sheet, i, 38));
//				dist.setQtde_dependentes(getValue(sheet, i, 39));
//				dist.setDt_admissao1(getValue(sheet, i, 40));
//				dist.setDtdemissao1(getValue(sheet, i, 41));
//				dist.setCpf1(getValue(sheet, i, 42));
//				dist.setValor_capital1(getValue(sheet, i, 43));
//				dist.setConstituicao1(getValue(sheet, i, 44));
//				dist.setPassivo_circulante(getValue(sheet, i, 45));
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
