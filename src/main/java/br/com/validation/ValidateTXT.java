package br.com.validation;

import java.io.BufferedReader;
import java.io.FileInputStream;
import java.io.InputStreamReader;
import java.util.ArrayList;
import java.util.List;

public class ValidateTXT extends OpenExcelValidate {

	int valor;

	public boolean validarTXT(String name, String caminhoPastaXML, String arquivoValidador) throws Exception {
//	public static void main(String[] args) throws Exception{	

		boolean conteudo = false;
		List<String> listaString = new ArrayList<String>();
		List<String> listaTemplate = new ArrayList<String>();
		boolean divergencia = false;
		String dado = "";

		try {
			// abertura do arquivo template
			BufferedReader myBuffer2 = new BufferedReader(
					new InputStreamReader(new FileInputStream(caminhoPastaXML + "template.txt"), "UTF-8"));

			String linha2 = myBuffer2.readLine();

			while (linha2 != null) {
				linha2 = myBuffer2.readLine();
				listaTemplate.add(linha2);
//				System.out.println(linha2);			
			}

//				String name2= "25200107526557001343550010000766471923862850-procNFe";
//				String name3= "21181256228356014868550030000000281413791763-procNFe";

			String name2 = name.replace(".xml", "");
			String path = arquivoValidador.replace("VALIDADOR.xlsb", "");

			// abertura do arquivo
			BufferedReader myBuffer = new BufferedReader(
					new InputStreamReader(new FileInputStream(path + name2 + ".txt"), "windows-1252"));

			// loop que lê e imprime todas as linhas do arquivo
			String linha = myBuffer.readLine();

			while (linha != null) {
				linha = myBuffer.readLine();
//			System.out.println(linha);

				if (linha.contains("SIMULAÇÃO NF")) {
					conteudo = false;
//    			System.out.println("===========================");
//    			for (int i=0; i<listaString.size(); i++) {
//    	            System.out.println(listaString.get(i));
//    	        }
//    			System.out.println("===========================");
					break;
				}

				if (conteudo == true) {
					listaString.add(linha);
				}

				if (linha.contains("DIFERENÇA")) {
//    			System.out.println("ENCONTRADO");
					conteudo = true;
				}

			}

			// Removendo últimas linhas em branco
			listaTemplate.remove(15);
			listaTemplate.remove(14);
			listaString.remove(14);
			listaString.remove(13);

			// Valida lista template com listas
			for (int i = 0; i < listaString.size(); i++) {
				dado = listaString.get(i);
//			num = i;
				if (listaTemplate.get(i).equals(dado)) {

//		    	System.out.println("Linha " + i + " sem divergência");
//		    	System.out.println("Esperado: " + listaTemplate.get(i));
//		    	System.out.println("Atual: " + listaString.get(i) );
					divergencia = false;

				} else {
//		    	System.out.println("Linha " + i + " apresenta divergência");
//		    	System.out.println("Esperado: " + listaTemplate.get(i));
//		    	System.out.println("Atual: " + listaString.get(i) );
//		    	System.out.println("APRESENTOU DIVERGENCIA");
					divergencia = true;
					break;
				}

			}

			myBuffer.close();
			return divergencia;

		} catch (Exception e) {
//			return true;
			System.out.println("Não foi possível validar o conteúdo do arquivo XML");
			System.out.println(e);
		}
		return true;
	}

	public int retornaContador() throws Exception {

//		int contadorUltimate = Integer.parseUnsignedInt(marcacao);
		return valor;
	}
}
