package br.com.validation;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.Scanner;
import java.util.Set;
import java.util.TreeMap;
import java.util.TreeSet;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.DefaultIndexedColorMap;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFormulaEvaluator;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelData {

	static ArrayList<String> nomeArquivos = new ArrayList();
	static String name;
	static String caminhoPastaXML;
	static List<String> listaString = new ArrayList<String>();
	static List<String> listaStringMateriais = new ArrayList<String>();
	static List<String> listaTemplate = new ArrayList<String>();
	static String divergencia = "Sem Divergencia";
	static boolean verify = false;
	static double newValue;
	static Object[] valor1Alterado;
	static Object[] valor2Alterado;
	static Object[] valor3Alterado;
	static Object[] valor4Alterado;
	static Object[] valor5Alterado;
	static Object[] valor6Alterado;
	static Object[] valor7Alterado;
	static Object[] valor8Alterado;
	static Object[] valor9Alterado;
	static Object[] valor10Alterado;
	static Object[] valor11Alterado;
	static Object[] valor12Alterado;
	static Object[] valor13Alterado;
	static Object[] valor14Alterado;

	// any exceptions need to be caught
//    public static void main(String[] args) throws Exception{
	public int writeExcel(String data, ArrayList<String> listaArquivosDivergencia, String arquivoValidador,
			String caminhoPasta) throws Exception {

//    List<String> listaString = new ArrayList<String>();
//    String numeracao="2";
		int numeracao = 2;
		int id = 0;
		Map<Integer, Object[]> reportData = new TreeMap<Integer, Object[]>();
		boolean conteudo = false;

//    	String divergencia = "Sem Divergencia";
		caminhoPastaXML = caminhoPasta;

		// workbook object
		XSSFWorkbook workbook = new XSSFWorkbook();

		// spreadsheet object
		XSSFSheet spreadsheet = workbook.createSheet("Report Execucao");
//        = workbook.createSheet(name);

		// creating a row object
		XSSFRow row;

		// CREATION
		String path = arquivoValidador.replace("VALIDADOR.xlsb", "");
		File diretorio = new File(path);

		for (File file : diretorio.listFiles()) {
			if (file.getName().contains("-procNFe.txt")) {
				listaString.clear();
				listaStringMateriais.clear();
				verify = false;
				nomeArquivos.add(file.getAbsolutePath());
				name = file.getName();

				// abertura do arquivo
				BufferedReader myBuffer = new BufferedReader(
						new InputStreamReader(new FileInputStream(path + name), "windows-1252"));

				// loop que lê e imprime todas as linhas do arquivo
				String linha = myBuffer.readLine();

				while (linha != null) {
					linha = myBuffer.readLine();
//					System.out.println(linha);

					if (linha.contains("SIMULAÇÃO NF")) {
						conteudo = false;
						break;
					}

					if (conteudo == true) {
						listaString.add(linha);
					}

					if (linha.contains("DIFERENÇA")) {
						conteudo = true;
					}

				}

				myBuffer.close();

				// ENCONTRA MATERIAIS
				// abertura do arquivo
				BufferedReader myBuffer2 = new BufferedReader(
						new InputStreamReader(new FileInputStream(path + name), "windows-1252"));

				// loop que lê e imprime todas as linhas do arquivo
				String linhaMaterial = myBuffer2.readLine();

				while (linhaMaterial != null) {
					linhaMaterial = myBuffer2.readLine();
//					System.out.println(linha);

					if (linhaMaterial.contains("Linha XML")) {
						conteudo = false;
						break;
					}

					if (conteudo == true) {
						listaStringMateriais.add(linhaMaterial);
					}

					if (linhaMaterial.contains("DETALHE NF")) {
						conteudo = true;
					}

				}

				myBuffer2.close();

				String materiaisTratados = (listaStringMateriais.get(1))
						.replace("DADOS DOS PRODUTOS\"	Código produto", "").trim();

				String[] listaMateriaisTratados = materiaisTratados.split("\t");

				String nameTratado = name.replace(".txt", ".xml");

				for (int n = 0; n < listaArquivosDivergencia.size(); n++) {
					if (listaArquivosDivergencia.get(n).contains(nameTratado)) {
						divergencia = "Com Divergencia";
						break;
					} else {

						divergencia = "Sem Divergencia";

					}
				}

				for (int n = 0; n < listaArquivosDivergencia.size(); n++) {
					if (listaArquivosDivergencia.get(n).contains(nameTratado)) {
						divergencia = "Com Divergencia";
						break;
					} else {
						divergencia = "Sem Divergencia";
					}
				}

				listaArquivosDivergencia.remove(nameTratado);

				preencheItens();

				String idxml = Integer.toString(id + 1);

				// ORG4 - LOGICA PARA IMPRIMIR .TXT NAO ENCONTRADOS

				// ORG5 - MONTA REPORTDATA COM DADOS
				for (int i = 0; i < 1; i++) {

					reportData.put(1, new Object[] { "Arquivo XML", "Status Geral", "Material",
							"Base de cálculo do ICMS PP (Próprio)", "Valor do ICMS PP (Próprio)",
							"Base de cálculo do FECOP PP (Próprio)", "Valor do FECOP PP (Próprio)",
							"Base de cálculo do ICMS ST", "Valor do ICMS ST", "Base de cálculo do FECOP ST",
							"Valor do FECOP ST", "Base de cálculo do IPI", "Valor do IPI", "Base de cálculo do PIS",
							"Valor do PIS", "Base de cálculo do COFINS", "Valor do COFINS" });

					// colocar FOR aqui para ver numero de itens

					for (int n = 0; n < listaMateriaisTratados.length; n++) {
						reportData.put(numeracao,
								new Object[] { name, divergencia, listaMateriaisTratados[n], valor1Alterado[n],
										valor2Alterado[n], valor3Alterado[n], valor4Alterado[n], valor5Alterado[n],
										valor6Alterado[n], valor7Alterado[n], valor8Alterado[n], valor9Alterado[n],
										valor10Alterado[n], valor11Alterado[n], valor12Alterado[n], valor13Alterado[n],
										valor14Alterado[n] });
						numeracao = numeracao + 1;
					}
					deleteFilesTXT(path);
				}
			}
		}

		// abertura do arquivo template
				BufferedReader myBuffer2 = new BufferedReader(
						new InputStreamReader(new FileInputStream(caminhoPastaXML + "template.txt"), "UTF-8"));

				listaString.clear();
				String linha2 = myBuffer2.readLine();

				while (linha2 != null) {
					linha2 = myBuffer2.readLine();
					listaString.add(linha2);
				}

		preencheItens();

		for (int y = 0; y < listaArquivosDivergencia.size(); y++) {
			System.out.println(listaArquivosDivergencia.get(y));
		}

		// ORG5 - MONTA REPORTDATA COM DADOS

		for (int i = 0; i < 1; i++) {

			reportData.put(1,
					new Object[] { "Arquivo XML", "Status Geral", "Material", "Base de cálculo do ICMS PP (Próprio)",
							"Valor do ICMS PP (Próprio)", "Base de cálculo do FECOP PP (Próprio)",
							"Valor do FECOP PP (Próprio)", "Base de cálculo do ICMS ST", "Valor do ICMS ST",
							"Base de cálculo do FECOP ST", "Valor do FECOP ST", "Base de cálculo do IPI",
							"Valor do IPI", "Base de cálculo do PIS", "Valor do PIS", "Base de cálculo do COFINS",
							"Valor do COFINS" });

			// colocar FOR aqui para ver numero de itens

			for (int n = 0; n < listaArquivosDivergencia.size(); n++) {
				reportData.put(numeracao,
						new Object[] { listaArquivosDivergencia.get(n), "Não foi possivel validar o XML", " ",
								valor1Alterado[n], valor2Alterado[n], valor3Alterado[n], valor4Alterado[n],
								valor5Alterado[n], valor6Alterado[n], valor7Alterado[n], valor8Alterado[n],
								valor9Alterado[n], valor10Alterado[n], valor11Alterado[n], valor12Alterado[n],
								valor13Alterado[n], valor14Alterado[n] });
				numeracao = numeracao + 1;
			}

//			deleteFilesTXT(path);

		}
		
		

		Set<Integer> keyid = new TreeSet<Integer>();
		keyid = reportData.keySet();

		int rowid = 0;

//        int rowid = count;

		// writing the data into the sheets...

		for (Integer key : keyid) {

			row = spreadsheet.createRow(rowid++);
			Object[] objectArr = reportData.get(key);
			int cellid = 0;

			for (Object obj : objectArr) {
				Cell cell = row.createCell(cellid++);
				cell.setCellValue((String) obj);

				if (key == 1) {
					XSSFCellStyle style;
					byte[] rgb;
					XSSFColor color;
					style = workbook.createCellStyle();
					style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
					rgb = new byte[3];
					rgb[0] = (byte) 211; // red
					rgb[1] = (byte) 211; // green
					rgb[2] = (byte) 211; // blue
					color = new XSSFColor(rgb, new DefaultIndexedColorMap());
					style.setFillForegroundColor(color);
					cell.setCellStyle(style);
				}

			}
		}

		// ATUALMENTE ENCERRA AQUI

//		System.out.println(divergencia);
		
		spreadsheet.autoSizeColumn(0);
		spreadsheet.autoSizeColumn(1);
		spreadsheet.autoSizeColumn(2);
		spreadsheet.autoSizeColumn(3);
		spreadsheet.autoSizeColumn(4);
		spreadsheet.autoSizeColumn(5);
		spreadsheet.autoSizeColumn(6);
		spreadsheet.autoSizeColumn(7);
		spreadsheet.autoSizeColumn(8);
		spreadsheet.autoSizeColumn(9);
		spreadsheet.autoSizeColumn(10);
		spreadsheet.autoSizeColumn(11);
		spreadsheet.autoSizeColumn(12);
		spreadsheet.autoSizeColumn(13);
		spreadsheet.autoSizeColumn(14);
		spreadsheet.autoSizeColumn(15);
		spreadsheet.autoSizeColumn(16);

		spreadsheet.setAutoFilter(CellRangeAddress.valueOf("A1:Q200"));

		//Compare XML
		XSSFSheet newSheet = workbook.cloneSheet(0, "Compare XML");
		Integer qtdRows = rowid;
		row = newSheet.createRow(rowid++);
		XSSFFormulaEvaluator formulaEvaluator = 
				  workbook.getCreationHelper().createFormulaEvaluator();
		Cell c = row.createCell(0);
		char alphabet = 'A';
		c.setCellValue("Comparação");
		for(int i = 1;i<=16;i++) {
			alphabet++;
			c = row.createCell(i);
			c.setCellFormula("IF("+alphabet+row.getRowNum()+"="+alphabet+(row.getRowNum()-1)+",\"OK\",\"Falha\")");
			formulaEvaluator.evaluateFormulaCell(c);
			CellStyle cellStyle = workbook.createCellStyle();
			switch (c.getStringCellValue()) {
			case "OK":
				cellStyle.setFillForegroundColor(IndexedColors.LIGHT_GREEN.getIndex());
			    cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
			    c.setCellStyle(cellStyle);
				break;
			case "Falha":
				cellStyle.setFillForegroundColor(IndexedColors.RED.getIndex());
			    cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
			    c.setCellStyle(cellStyle);
				break;
			default:
				System.out.println("Não foi encontrado a opção certa para pintar a célula.");
				break;
			}
		}
		
		// .xlsx is the format for Excel Sheets...
		// writing the workbook into the file...
		FileOutputStream out = new FileOutputStream(new File(path + "/report/report.xlsx"));

		workbook.write(out);

		// copiar report
		saveBackupReport(path);

		out.close();

//		}

		return rowid;
	}

	public static void preencheItens() {
		// ORG3 - COLOCA VALOR NOS ITENS

		String valor1 = (listaString.get(0)).replace("Base de cálculo do ICMS PP (Próprio)", "");
		String valorVerificaArred1 = valor1.trim();
		String valor1uptd = valor1.replace("\t\t\t\t", "");
		valor1Alterado = valor1uptd.split("\t ");

		if (!valorVerificaArred1.isEmpty() && !valorVerificaArred1.contains("\t")) {
			newValue = converteValor(valorVerificaArred1);
			if (verify == false) {
				verify = verificaArredondamento(newValue);
			}
		} else {
			newValue = 1.0;
			if (!valorVerificaArred1.isEmpty()) {
			verify = true;
			divergencia = "Com Divergencia";
			}
		}

//	if(newValue<0.5) {
//		valor1Alterado = valor1.replace(valor1, "");
//	}

		String valor2 = (listaString.get(1)).replace("Valor do ICMS PP (Próprio)", "");
		String valorVerificaArred2 = valor2.trim();
		String valor2uptd = valor2.replace("\t\t\t\t", "");
		valor2Alterado = valor2uptd.split("\t ");

		if (!valorVerificaArred2.isEmpty() && !valorVerificaArred2.contains("\t")) {
			newValue = converteValor(valorVerificaArred2);
			if (verify == false) {
				verify = verificaArredondamento(newValue);
			}
		} else {
			newValue = 1.0;
			if (!valorVerificaArred2.isEmpty()) {
				verify = true;
				divergencia = "Com Divergencia";
				}
		}

//	if(newValue<0.5) {
//		valor2Alterado = valor2.replace(valor2, "");
//	}
		String valor3 = (listaString.get(2)).replace("Base de cálculo do FECOP PP (Próprio)", "");
		String valorVerificaArred3 = valor3.trim();
		String valor3uptd = valor3.replace("\t\t\t\t", "");
		valor3Alterado = valor3uptd.split("\t ");

		if (!valorVerificaArred3.isEmpty() && !valorVerificaArred3.contains("\t")) {
			newValue = converteValor(valorVerificaArred3);
			if (verify == false) {
				verify = verificaArredondamento(newValue);
			}
		} else {
			newValue = 1.0;
			if (!valorVerificaArred3.isEmpty()) {
				verify = true;
				divergencia = "Com Divergencia";
				}
		}
//
//	if(newValue<0.5) {
//		valor3Alterado = valor3.replace(valor3, "");
//	}

		String valor4 = (listaString.get(3)).replace("Valor do FECOP PP (Próprio)", "");
		String valorVerificaArred4 = valor4.trim();
		String valor4uptd = valor4.replace("\t\t\t\t", "");
		valor4Alterado = valor4uptd.split("\t ");

		if (!valorVerificaArred4.isEmpty() && !valorVerificaArred4.contains("\t")) {
			newValue = converteValor(valorVerificaArred4);
			if (verify == false) {
				verify = verificaArredondamento(newValue);
			}
		} else {
			newValue = 1.0;
			if (!valorVerificaArred4.isEmpty()) {
				verify = true;
				divergencia = "Com Divergencia";
				}
		}

//	if(newValue<0.5) {
//		valor4Alterado = valor4.replace(valor4, "");
//	}
		String valor5 = (listaString.get(4)).replace("Base de cálculo do ICMS ST", "");
		String valorVerificaArred5 = valor5.trim();
		String valor5uptd = valor5.replace("\t\t\t\t", "");
		valor5Alterado = valor5uptd.split("\t ");

//	for(int i=0; i<valor5Split.length; i++)
//	{
//	System.out.println("Linha: " + i + " e valor é: " + valor5Split[i]);
//	}

		if (!valorVerificaArred5.isEmpty() && !valorVerificaArred5.contains("\t")) {
			newValue = converteValor(valorVerificaArred5);
			if (verify == false) {
				verify = verificaArredondamento(newValue);
			}
		} else {
			newValue = 1.0;
			if (!valorVerificaArred5.isEmpty()) {
				verify = true;
				divergencia = "Com Divergencia";
				}
		}
//
//	if(newValue<0.5) {
//		valor5Alterado = valor5.replace(valor5, "");
//	}

		String valor6 = (listaString.get(5)).replace("Valor do ICMS ST", "");
		String valorVerificaArred6 = valor6.trim();
		String valor6uptd = valor6.replace("\t\t\t\t", "");
		valor6Alterado = valor6uptd.split("\t ");

		if (!valorVerificaArred6.isEmpty() && !valorVerificaArred6.contains("\t")) {
			newValue = converteValor(valorVerificaArred6);
			if (verify == false) {
				verify = verificaArredondamento(newValue);
			}
		} else {
			newValue = 1.0;
			if (!valorVerificaArred6.isEmpty()) {
				verify = true;
				divergencia = "Com Divergencia";
				}
		}

//	if(newValue<0.5) {
//		valor6Alterado = valor6.replace(valor6, "");
//	}		

		String valor7 = (listaString.get(6)).replace("Base de cálculo do FECOP ST", "");
		String valorVerificaArred7 = valor7.trim();
		String valor7uptd = valor7.replace("\t\t\t\t", "");
		valor7Alterado = valor7uptd.split("\t ");

		if (!valorVerificaArred7.isEmpty() && !valorVerificaArred7.contains("\t")) {
			newValue = converteValor(valorVerificaArred7);
			if (verify == false) {
				verify = verificaArredondamento(newValue);
			}
		} else {
			newValue = 1.0;
			if (!valorVerificaArred7.isEmpty()) {
				verify = true;
				divergencia = "Com Divergencia";
				}
		}

//	if(newValue<0.5) {
//		valor7Alterado = valor7.replace(valor7, "");
//	}

		String valor8 = (listaString.get(7)).replace("Valor do FECOP ST", "");
		String valorVerificaArred8 = valor8.trim();
		String valor8uptd = valor8.replace("\t\t\t\t", "");
		valor8Alterado = valor8uptd.split("\t ");

		if (!valorVerificaArred8.isEmpty() && !valorVerificaArred8.contains("\t")) {
			newValue = converteValor(valorVerificaArred8);
			if (verify == false) {
				verify = verificaArredondamento(newValue);
			}
		} else {
			newValue = 1.0;
			if (!valorVerificaArred8.isEmpty()) {
				verify = true;
				divergencia = "Com Divergencia";
				}
		}

//	if(newValue<0.5) {
//		valor8Alterado = valor8.replace(valor8, "");
//	}

		String valor9 = (listaString.get(8)).replace("Base de cálculo do IPI", "");
		String valorVerificaArred9 = valor9.trim();
		String valor9uptd = valor9.replace("\t\t\t\t", "");
		valor9Alterado = valor9uptd.split("\t ");

		if (!valorVerificaArred9.isEmpty() && !valorVerificaArred9.contains("\t")) {
			newValue = converteValor(valorVerificaArred9);
			if (verify == false) {
				verify = verificaArredondamento(newValue);
			}
		} else {
			newValue = 1.0;
			if (!valorVerificaArred9.isEmpty()) {
				verify = true;
				divergencia = "Com Divergencia";
				}
		}

//	if(newValue<0.5) {
//		valor9Alterado = valor9.replace(valor9, "");
//	}

		String valor10 = (listaString.get(9)).replace("Valor do IPI", "");
		String valorVerificaArred10 = valor10.trim();
		String valor10uptd = valor10.replace("\t\t\t\t", "");
		valor10Alterado = valor10uptd.split("\t ");

		if (!valorVerificaArred10.isEmpty() && !valorVerificaArred10.contains("\t")) {
			newValue = converteValor(valorVerificaArred10);
			if (verify == false) {
				verify = verificaArredondamento(newValue);
			}
		} else {
			newValue = 1.0;
			if (!valorVerificaArred10.isEmpty()) {
				verify = true;
				divergencia = "Com Divergencia";
				}
		}

//	if(newValue<0.5) {
//		valor10Alterado = valor10.replace(valor10, "");
//	}

		String valor11 = (listaString.get(10)).replace("Base de cálculo do PIS", "");
		String valorVerificaArred11 = valor11.trim();
		String valor11uptd = valor11.replace("\t\t\t\t", "");
		valor11Alterado = valor11uptd.split("\t ");

		if (!valorVerificaArred11.isEmpty() && !valorVerificaArred11.contains("\t")) {
			newValue = converteValor(valorVerificaArred11);
			if (verify == false) {
				verify = verificaArredondamento(newValue);
			}
		} else {
			newValue = 1.0;
			if (!valorVerificaArred11.isEmpty()) {
				verify = true;
				divergencia = "Com Divergencia";
				}
		}

//	if(newValue<0.5) {
//		valor11Alterado = valor11.replace(valor11, "");
//	}

		String valor12 = (listaString.get(11)).replace("Valor do PIS", "");
		String valorVerificaArred12 = valor12.trim();
		String valor12uptd = valor12.replace("\t\t\t\t", "");
		valor12Alterado = valor12uptd.split("\t ");

		if (!valorVerificaArred12.isEmpty() && !valorVerificaArred12.contains("\t")) {
			newValue = converteValor(valorVerificaArred12);
			if (verify == false) {
				verify = verificaArredondamento(newValue);
			}
		} else {
			newValue = 1.0;
			if (!valorVerificaArred12.isEmpty()) {
				verify = true;
				divergencia = "Com Divergencia";
				}
		}

//	if(newValue<0.5) {
//		valor12Alterado = valor12.replace(valor12, "");
//	}

		String valor13 = (listaString.get(12)).replace("Base de cálculo do COFINS", "");
		String valorVerificaArred13 = valor13.trim();
		String valor13uptd = valor13.replace("\t\t\t\t", "");
		valor13Alterado = valor13uptd.split("\t ");

		if (!valorVerificaArred13.isEmpty() && !valorVerificaArred13.contains("\t")) {
			newValue = converteValor(valorVerificaArred13);
			if (verify == false) {
				verify = verificaArredondamento(newValue);
			}
		} else {
			newValue = 1.0;
			if (!valorVerificaArred13.isEmpty()) {
				verify = true;
				divergencia = "Com Divergencia";
				}
		}

//	if(newValue<0.5) {
//		valor13Alterado = valor13.replace(valor13, "");
//	}

		String valor14 = (listaString.get(13)).replace("Valor do COFINS", "");
		String valorVerificaArred14 = valor14.trim();
		String valor14uptd = valor14.replace("\t\t\t\t", "");
		valor14Alterado = valor14uptd.split("\t ");

		if (!valorVerificaArred14.isEmpty() && !valorVerificaArred14.contains("\t")) {
			newValue = converteValor(valorVerificaArred14);
			if (verify == false) {
				verify = verificaArredondamento(newValue);
			}
		} else {
			newValue = 1.0;
			if (!valorVerificaArred14.isEmpty()) {
				verify = true;
				divergencia = "Com Divergencia";
				}
		}

//	if(newValue<0.5) {
//		valor14Alterado = valor14.replace(valor14, "");
//	}	      

	}
	public static void deleteFilesTXT(String path) {
		String nameFileTxt = path + name;
		File f1 = new File(nameFileTxt);
		f1.delete();
	}

	public static boolean verificaArredondamento(double valorDouble) {
		if (valorDouble > 0.5) {
			divergencia = "Com Divergencia";
			return true;
		} else {
			divergencia = "Ressalva (Arrendondamento)";
		}
		return false;
	}

	public static double converteValor(String valor) {
		if (valor.contains("#VALUE!") || valor.contains("N/A") || valor.contains("#VALOR!") || valor.contains("N/D")) {
			return 9999999;
		} else {
			String valorAlterado = valor.replace("(", "").trim();
			String valorAlterado2 = valorAlterado.replace(")", "").trim();
			String valorAlterado3 = valorAlterado2.replace(".", "");
			String valorAlterado4 = valorAlterado3.replace(",", ".");
			Double valorDouble = Double.parseDouble(valorAlterado4);
			return valorDouble;
		}
	}

	public static void saveBackupReport(String path) {
		String nameFile = path + "report\\report.xlsx";

		try {
			String[] command = new String[5];
			command[0] = "cmd";
			command[1] = "/c";
			command[2] = "copy";
			command[3] = nameFile;
			command[4] = caminhoPastaXML;
			Process p = Runtime.getRuntime().exec(command);

			InputStream in = p.getInputStream();
			Scanner scan = new Scanner(in);
			while (scan.hasNext()) {
				System.out.println(scan.nextLine());
			}

		} catch (IOException e) {
			e.printStackTrace();
		}
	}

}
