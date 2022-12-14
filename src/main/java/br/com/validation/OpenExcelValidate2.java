package br.com.validation;

import java.awt.Robot;
import java.awt.Toolkit;
import java.awt.datatransfer.StringSelection;
import java.awt.event.KeyEvent;
import java.io.BufferedReader;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.Map;
import java.util.Scanner;
import java.util.TreeMap;

import javax.validation.constraints.AssertTrue;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.util.ZipSecureFile;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.util.IOUtils;


public class OpenExcelValidate2 {

	final static String nome_planilha_xml = "planilhaXML.xlsx";
	final static String nome_report = "report.xlsx";
	static String name;
	static String caminhoPastaXML;
	static String arquivoValidador;
	static ArrayList<String> nomeArquivos = new ArrayList();
	static ArrayList<String> listaArquivosDivergencia = new ArrayList();

	static final Runtime run = Runtime.getRuntime();
	static Process pro;
	static BufferedReader read;

	static String reportName;
	static String reportPath;
	static String reportPathDivergencia;
	static String path;
	static int num = 0;
	static String data;


	public static void main(String[] args) throws Exception {

		for (int i = 0; i < args.length; i++) {
			System.out.println(args[i]);
		}

		try {
			caminhoPastaXML = (args[0]);
			arquivoValidador = (args[1]);

			File diretorio = new File(caminhoPastaXML);

			data = new SimpleDateFormat("yyyyMMddhhmm").format(new Date());

			reportName = "report_" + data;
			reportPath = caminhoPastaXML + reportName;
			reportPathDivergencia = reportPath + "\\DIVERGENCIAS";
			new File(reportPath).mkdir();
			new File(reportPathDivergencia).mkdir();

			path = arquivoValidador.replace("VALIDADOR.xlsb", "");
			//      name = path + "23190507526557000886550010001439911773469596-procNFe.xml";

			ArrayList<String> nomesArquivosXML = new ArrayList<String>();
/*
			//CRIAR XML REPORT
			for (File file : diretorio.listFiles()) {
				if (file.getName().contains(".xml")) {
					nomeArquivos.add(file.getAbsolutePath());
					name = file.getName();
					System.out.println(name);
					Thread.sleep(5000);

					// Abre arquivo Validador e espera carregar - PARAMETER2
					java.awt.Desktop.getDesktop().open(new File(arquivoValidador));
					Thread.sleep(50000);

					// Importar arquivo
					Robot robot = new Robot();
					robot.keyPress(KeyEvent.VK_ALT);
					robot.keyPress(KeyEvent.VK_X);
					Thread.sleep(3000);
					uploadFile(file.getAbsolutePath());
					Thread.sleep(20000);
					robot.keyPress(KeyEvent.VK_ALT);
					Thread.sleep(2500);
					robot.keyPress(KeyEvent.VK_S);
					Thread.sleep(1000);
					robot.keyPress(KeyEvent.VK_ALT);
					robot.keyRelease(KeyEvent.VK_ALT);
					Thread.sleep(3000);
					
					for (int i = 0; i < 20; i++) {
						// reexibir planilhas
						robot.keyPress(KeyEvent.VK_ALT);
						robot.keyRelease(KeyEvent.VK_ALT);
						robot.keyPress(KeyEvent.VK_C);
						robot.keyRelease(KeyEvent.VK_C);
						robot.keyPress(KeyEvent.VK_O);
						robot.keyRelease(KeyEvent.VK_O);
						robot.keyPress(KeyEvent.VK_U);
						robot.keyRelease(KeyEvent.VK_U);
						robot.keyPress(KeyEvent.VK_I);
						robot.keyRelease(KeyEvent.VK_I);
						Thread.sleep(2000);
						robot.keyPress(KeyEvent.VK_DOWN);
						robot.keyRelease(KeyEvent.VK_DOWN);
						robot.keyPress(KeyEvent.VK_DOWN);
						robot.keyRelease(KeyEvent.VK_DOWN);
						robot.keyPress(KeyEvent.VK_DOWN);
						robot.keyRelease(KeyEvent.VK_DOWN);
						robot.keyPress(KeyEvent.VK_ENTER);
						robot.keyRelease(KeyEvent.VK_ENTER);
						Thread.sleep(1000);
					}
					
					//Deletar as 20 planilhas abertas
					for (int i = 0; i < 20; i++) {
						// deletar planilha
						robot.keyPress(KeyEvent.VK_ALT);
						robot.keyRelease(KeyEvent.VK_ALT);
						robot.keyPress(KeyEvent.VK_C);
						robot.keyRelease(KeyEvent.VK_C);
						robot.keyPress(KeyEvent.VK_K);
						robot.keyRelease(KeyEvent.VK_K);
						robot.keyPress(KeyEvent.VK_E);
						robot.keyRelease(KeyEvent.VK_E);
						Thread.sleep(500);
						robot.keyPress(KeyEvent.VK_ENTER);
						Thread.sleep(500);
					}
					
					for (int i = 0; i < 2; i++) {
						// reexibir planilhas
						robot.keyPress(KeyEvent.VK_ALT);
						robot.keyRelease(KeyEvent.VK_ALT);
						robot.keyPress(KeyEvent.VK_C);
						robot.keyRelease(KeyEvent.VK_C);
						robot.keyPress(KeyEvent.VK_O);
						robot.keyRelease(KeyEvent.VK_O);
						robot.keyPress(KeyEvent.VK_U);
						robot.keyRelease(KeyEvent.VK_U);
						robot.keyPress(KeyEvent.VK_I);
						robot.keyRelease(KeyEvent.VK_I);
						Thread.sleep(2000);
						robot.keyPress(KeyEvent.VK_ENTER);
						robot.keyRelease(KeyEvent.VK_ENTER);
						Thread.sleep(1000);
					}
					
					//Deletar as 3 planilhas abertas
					for (int i = 0; i < 2; i++) {
						// deletar planilha
						robot.keyPress(KeyEvent.VK_ALT);
						robot.keyRelease(KeyEvent.VK_ALT);
						robot.keyPress(KeyEvent.VK_C);
						robot.keyRelease(KeyEvent.VK_C);
						robot.keyPress(KeyEvent.VK_K);
						robot.keyRelease(KeyEvent.VK_K);
						robot.keyPress(KeyEvent.VK_E);
						robot.keyRelease(KeyEvent.VK_E);
						Thread.sleep(500);
						robot.keyPress(KeyEvent.VK_ENTER);
						Thread.sleep(500);
					}
					
					// reexibir planilhas
					robot.keyPress(KeyEvent.VK_ALT);
					robot.keyRelease(KeyEvent.VK_ALT);
					robot.keyPress(KeyEvent.VK_C);
					robot.keyRelease(KeyEvent.VK_C);
					robot.keyPress(KeyEvent.VK_O);
					robot.keyRelease(KeyEvent.VK_O);
					robot.keyPress(KeyEvent.VK_U);
					robot.keyRelease(KeyEvent.VK_U);
					robot.keyPress(KeyEvent.VK_I);
					robot.keyRelease(KeyEvent.VK_I);
					Thread.sleep(2000);
					robot.keyPress(KeyEvent.VK_ENTER);
					robot.keyRelease(KeyEvent.VK_ENTER);
					Thread.sleep(1000);
					
					
					// mudar planilha
					robot.keyPress(KeyEvent.VK_CONTROL);
					robot.keyPress(KeyEvent.VK_PAGE_UP);
					robot.keyRelease(KeyEvent.VK_PAGE_UP);
					robot.keyRelease(KeyEvent.VK_CONTROL);


					// deletar planilha
					robot.keyPress(KeyEvent.VK_ALT);
					robot.keyRelease(KeyEvent.VK_ALT);
					robot.keyPress(KeyEvent.VK_C);
					robot.keyRelease(KeyEvent.VK_C);
					robot.keyPress(KeyEvent.VK_K);
					robot.keyRelease(KeyEvent.VK_K);
					robot.keyPress(KeyEvent.VK_E);
					robot.keyRelease(KeyEvent.VK_E);
					Thread.sleep(500);
					robot.keyPress(KeyEvent.VK_ENTER);
					
					
					// Salvar arquivo validado
					robot.keyPress(KeyEvent.VK_F12);
					robot.keyRelease(KeyEvent.VK_F12);
					Thread.sleep(3000);
					// criar pasta
					System.out.println(reportPath+name);
					setClipboardData("lerXML_"+name);
					nomesArquivosXML.add("lerXML_"+name);
					Thread.sleep(1000);
					
					
					robot.keyPress(KeyEvent.VK_CONTROL);
					robot.keyPress(KeyEvent.VK_V);
					robot.keyRelease(KeyEvent.VK_V);
					robot.keyRelease(KeyEvent.VK_CONTROL);
					//selecionar tipo do arquivo
					robot.keyPress(KeyEvent.VK_TAB);
					robot.keyRelease(KeyEvent.VK_TAB);
					System.out.println("Selecionar tipo Arquivo");
					for(int i = 0;i<9;i++) {
						Thread.sleep(1000);
					robot.keyPress(KeyEvent.VK_P);
					robot.keyRelease(KeyEvent.VK_P);
					}
					robot.keyPress(KeyEvent.VK_ENTER);
					robot.keyRelease(KeyEvent.VK_ENTER);
					Thread.sleep(1500);
					robot.keyPress(KeyEvent.VK_ENTER);
					robot.keyRelease(KeyEvent.VK_ENTER);
					Thread.sleep(3000);

					// Fechar arquivo
					robot.keyPress(KeyEvent.VK_ALT);
					robot.keyPress(KeyEvent.VK_F4);
					robot.keyRelease(KeyEvent.VK_F4);
					robot.keyRelease(KeyEvent.VK_ALT);
					Thread.sleep(3000);
					robot.keyPress(KeyEvent.VK_ENTER);
					robot.keyRelease(KeyEvent.VK_ENTER);
					Thread.sleep(1500);

				}
			}

			//CRIAR ARQUIVO MESCLADO DOS XMLS
			// workbook object
			XSSFWorkbook workbook = new XSSFWorkbook();

			// spreadsheet object
			XSSFSheet destinationWorksheet = workbook.createSheet("Dados XML");
			//	        = workbook.createSheet(name);

			// CREATION
			String path = arquivoValidador.replace("VALIDADOR.xlsb", "");
			diretorio = new File(path);
			IOUtils.setByteArrayMaxOverride(500000000);
			int contadorXML = 0;
			int contadorDestino = 0;
			int contadorSource = 0;
			
			//Armazenar as Planilhas que contem as informações do XML
			ArrayList<XSSFWorkbook> workbooks = new ArrayList<XSSFWorkbook>(); 
			for (File file : diretorio.listFiles()) {	
				if (file.getName().contains("lerXML_")) {
					System.out.println(file.getAbsolutePath());
					workbooks.add(new XSSFWorkbook(new FileInputStream(file.getAbsolutePath())));
				}
			}
			
			int contadorWorkbook = workbooks.size();
			
			//Inicio Cabeçalho
			XSSFSheet tempSheet1 = workbooks.get(1).getSheet("LerXml");
			for(int contadorLinha = 0;contadorLinha<2;contadorLinha++) {
				// Get the source / new row
				XSSFRow sourceRow = tempSheet1.getRow(contadorLinha);
				XSSFRow newRow = destinationWorksheet.createRow(contadorDestino);
				// If the old cell is null jump to next cell
				if (isRowEmpty(sourceRow)) {
					continue;
				}

				// Loop through source columns to add to new row
				for (int j = 0; j < sourceRow.getLastCellNum(); j++) {
					// Grab a copy of the old/new cell
					XSSFCell oldCell = sourceRow.getCell(j);
					XSSFCell newCell = newRow.createCell(j);
					// If the old cell is null jump to next cell
					if (oldCell == null) {
						continue;
					}

					// Copy style from old cell and apply to new cell
					XSSFCellStyle newCellStyle = workbook.createCellStyle();
					newCellStyle.cloneStyleFrom(oldCell.getCellStyle());
					newCell.setCellStyle(newCellStyle);

					// If there is a cell comment, copy
					if (oldCell.getCellComment() != null) {
						newCell.setCellComment(oldCell.getCellComment());
					}

					// If there is a cell hyperlink, copy
					if (oldCell.getHyperlink() != null) {
						newCell.setHyperlink(oldCell.getHyperlink());
					}

					// Set the cell data type
					newCell.setCellType(oldCell.getCellType());

					// Set the cell data value
					switch (oldCell.getCellType()) {
					case BLANK:
						newCell.setCellValue(oldCell.getStringCellValue());
						break;
					case BOOLEAN:
						newCell.setCellValue(oldCell.getBooleanCellValue());
						break;
					case ERROR:
						newCell.setCellErrorValue(oldCell.getErrorCellValue());
						break;
					case FORMULA:
						newCell.setCellFormula(oldCell.getCellFormula());
						break;
					case NUMERIC:
						newCell.setCellValue(oldCell.getNumericCellValue());
						break;
					case STRING:
						newCell.setCellValue(oldCell.getRichStringCellValue());
						break;
					}
				}
					// If there are are any merged regions in the source row, copy to new row
					for (int h = 0; h < tempSheet1.getNumMergedRegions(); h++) {
						CellRangeAddress cellRangeAddress = tempSheet1.getMergedRegion(h);
						if (cellRangeAddress.getFirstRow() == sourceRow.getRowNum()) {
							CellRangeAddress newCellRangeAddress = new CellRangeAddress(newRow.getRowNum(),
									(newRow.getRowNum() + (cellRangeAddress.getLastRow() - cellRangeAddress.getFirstRow())),
									cellRangeAddress.getFirstColumn(), cellRangeAddress.getLastColumn());
							destinationWorksheet.addMergedRegion(newCellRangeAddress);
						}
					}
					contadorDestino++;
			}
			//Fim Cabeçalho

			//Inicio Corpo XML para Planilha
			for (int i = 0; i < contadorWorkbook; i++) {
				XSSFSheet tempSheet = workbooks.get(i).getSheet("LerXml");
				for (int j = 0; j < 2; j++) {
						XSSFRow sourceRow = tempSheet.getRow(j);
						tempSheet.removeRow(sourceRow);
				}
				if(i==0) {
					contadorDestino = 2;
				}
				if(i==1) {
					contadorDestino = 3;
				}
				for(int contadorLinha = 0;contadorLinha<tempSheet.getPhysicalNumberOfRows();contadorLinha++) {
					// Get the source / new row
					XSSFRow sourceRow = tempSheet.getRow(contadorLinha);
					XSSFRow newRow = destinationWorksheet.createRow(contadorDestino);
					// If the old cell is null jump to next cell
					if (isRowEmpty(sourceRow)) {
						continue;
					}

					// Loop through source columns to add to new row
					for (int j = 0; j < sourceRow.getLastCellNum(); j++) {
						// Grab a copy of the old/new cell
						XSSFCell oldCell = sourceRow.getCell(j);
						XSSFCell newCell = newRow.createCell(j);
						// If the old cell is null jump to next cell
						if (oldCell == null) {
							continue;
						}

						// Copy style from old cell and apply to new cell
						XSSFCellStyle newCellStyle = workbook.createCellStyle();
						newCellStyle.cloneStyleFrom(oldCell.getCellStyle());
						newCell.setCellStyle(newCellStyle);

						// If there is a cell comment, copy
						if (oldCell.getCellComment() != null) {
							newCell.setCellComment(oldCell.getCellComment());
						}

						// If there is a cell hyperlink, copy
						if (oldCell.getHyperlink() != null) {
							newCell.setHyperlink(oldCell.getHyperlink());
						}

						// Set the cell data type
						newCell.setCellType(oldCell.getCellType());

						// Set the cell data value
						switch (oldCell.getCellType()) {
						case BLANK:
							newCell.setCellValue(oldCell.getStringCellValue());
							break;
						case BOOLEAN:
							newCell.setCellValue(oldCell.getBooleanCellValue());
							break;
						case ERROR:
							newCell.setCellErrorValue(oldCell.getErrorCellValue());
							break;
						case FORMULA:
							newCell.setCellFormula(oldCell.getCellFormula());
							break;
						case NUMERIC:
							newCell.setCellValue(oldCell.getNumericCellValue());
							break;
						case STRING:
							newCell.setCellValue(oldCell.getRichStringCellValue());
							break;
						}
					}
					if(i<1) {
						// If there are are any merged regions in the source row, copy to new row
						for (int h = 0; h < tempSheet.getNumMergedRegions(); h++) {
							CellRangeAddress cellRangeAddress = tempSheet.getMergedRegion(h);
							if (cellRangeAddress.getFirstRow() == sourceRow.getRowNum()) {
								CellRangeAddress newCellRangeAddress = new CellRangeAddress(newRow.getRowNum(),
										(newRow.getRowNum() + (cellRangeAddress.getLastRow() - cellRangeAddress.getFirstRow())),
										cellRangeAddress.getFirstColumn(), cellRangeAddress.getLastColumn());
								destinationWorksheet.addMergedRegion(newCellRangeAddress);
							}
						}
					}
					contadorDestino=contadorDestino+3;
				}
			}
			//Fim Corpo XML para Planilha
			
			//Remover/Fechar planilhas abertar
			for (XSSFWorkbook xssfWorkbook : workbooks) {
				xssfWorkbook.close();
			}
			
			// .xlsx is the format for Excel Sheets...
			// writing the workbook into the file...
			FileOutputStream out = new FileOutputStream(new File(caminhoPastaXML+"nome_planilha_xml"));

			workbook.write(out);

			out.close();
			workbook.close();

			

			//INICIO AJUSTAR LAYOUT REPORT
			diretorio = new File(path + "/report");
			XSSFWorkbook workbookReport = null;
			File fileReport = null;
			for (File file : diretorio.listFiles()) {
				System.out.println(file.getAbsolutePath());
				if (file.getName().contains("nome_report")) {
					System.out.println("dentro if:" +file.getAbsolutePath());
						fileReport = file;
						workbookReport = new XSSFWorkbook(new FileInputStream(file.getAbsolutePath()));
						break;
				}
			}
			
			XSSFSheet sheet = workbookReport.getSheetAt(0);
			XSSFSheet sheetTemp = workbookReport.createSheet("Report Final");

			contadorDestino = 0;
			//Inicio Cabeçalho
			tempSheet1 = workbookReport.getSheetAt(0);
			for(int contadorLinha = 0;contadorLinha<2;contadorLinha++) {
				// Get the source / new row
				XSSFRow sourceRow = tempSheet1.getRow(contadorLinha);
				XSSFRow newRow = sheetTemp.createRow(contadorDestino);
				// If the old cell is null jump to next cell
				if (isRowEmpty(sourceRow)) {
					continue;
				}

				// Loop through source columns to add to new row
				for (int j = 0; j < sourceRow.getLastCellNum(); j++) {
					// Grab a copy of the old/new cell
					XSSFCell oldCell = sourceRow.getCell(j);
					XSSFCell newCell = newRow.createCell(j);
					// If the old cell is null jump to next cell
					if (oldCell == null) {
						continue;
					}

					// Copy style from old cell and apply to new cell
					XSSFCellStyle newCellStyle = workbookReport.createCellStyle();
					newCellStyle.cloneStyleFrom(oldCell.getCellStyle());
					newCell.setCellStyle(newCellStyle);

					// If there is a cell comment, copy
					if (oldCell.getCellComment() != null) {
						newCell.setCellComment(oldCell.getCellComment());
					}

					// If there is a cell hyperlink, copy
					if (oldCell.getHyperlink() != null) {
						newCell.setHyperlink(oldCell.getHyperlink());
					}

					// Set the cell data type
					newCell.setCellType(oldCell.getCellType());

					// Set the cell data value
					switch (oldCell.getCellType()) {
					case BLANK:
						newCell.setCellValue(oldCell.getStringCellValue());
						break;
					case BOOLEAN:
						newCell.setCellValue(oldCell.getBooleanCellValue());
						break;
					case ERROR:
						newCell.setCellErrorValue(oldCell.getErrorCellValue());
						break;
					case FORMULA:
						newCell.setCellFormula(oldCell.getCellFormula());
						break;
					case NUMERIC:
						newCell.setCellValue(oldCell.getNumericCellValue());
						break;
					case STRING:
						newCell.setCellValue(oldCell.getRichStringCellValue());
						break;
					}
				}
					// If there are are any merged regions in the source row, copy to new row
					for (int h = 0; h < tempSheet1.getNumMergedRegions(); h++) {
						CellRangeAddress cellRangeAddress = tempSheet1.getMergedRegion(h);
						if (cellRangeAddress.getFirstRow() == sourceRow.getRowNum()) {
							CellRangeAddress newCellRangeAddress = new CellRangeAddress(newRow.getRowNum(),
									(newRow.getRowNum() + (cellRangeAddress.getLastRow() - cellRangeAddress.getFirstRow())),
									cellRangeAddress.getFirstColumn(), cellRangeAddress.getLastColumn());
							sheetTemp.addMergedRegion(newCellRangeAddress);
						}
					}
					contadorDestino++;
			}
			//Fim Cabeçalho
			
			int contador = 1;
			
			//Inicio Corpo XML para Planilha
			for (int j = 0; j < 1; j++) {
				XSSFRow sourceRow = sheet.getRow(j);
				sheet.removeRow(sourceRow);
			}
			int numberofrows=sheet.getPhysicalNumberOfRows();
			for (int i = 0; i < 2; i++) {

				if(i==0) {
					contador = 1;
					contadorDestino = 1;
				}
				if(i==1) {
					contadorDestino = 2;
				}
				for(int contadorLinha = 0;contadorLinha<=numberofrows;contadorLinha++) {
					// Get the source / new row
					XSSFRow sourceRow = sheet.getRow(contadorLinha);
					XSSFRow newRow = sheetTemp.createRow(contadorDestino);
					// If the old cell is null jump to next cell
					if (isRowEmpty(sourceRow)) {
						continue;
					}

					// Loop through source columns to add to new row
					for (int j = 0; j < sourceRow.getLastCellNum(); j++) {
						// Grab a copy of the old/new cell
						XSSFCell oldCell = sourceRow.getCell(j);
						XSSFCell newCell = newRow.createCell(j);
						// If the old cell is null jump to next cell
						if (oldCell == null) {
							continue;
						}

						// Copy style from old cell and apply to new cell
						XSSFCellStyle newCellStyle = workbookReport.createCellStyle();
						newCellStyle.cloneStyleFrom(oldCell.getCellStyle());
						newCell.setCellStyle(newCellStyle);

						// If there is a cell comment, copy
						if (oldCell.getCellComment() != null) {
							newCell.setCellComment(oldCell.getCellComment());
						}

						// If there is a cell hyperlink, copy
						if (oldCell.getHyperlink() != null) {
							newCell.setHyperlink(oldCell.getHyperlink());
						}

						// Set the cell data type
						newCell.setCellType(oldCell.getCellType());

						// Set the cell data value
						switch (oldCell.getCellType()) {
						case BLANK:
							newCell.setCellValue(oldCell.getStringCellValue());
							break;
						case BOOLEAN:
							newCell.setCellValue(oldCell.getBooleanCellValue());
							break;
						case ERROR:
							newCell.setCellErrorValue(oldCell.getErrorCellValue());
							break;
						case FORMULA:
							newCell.setCellFormula(oldCell.getCellFormula());
							break;
						case NUMERIC:
							newCell.setCellValue(oldCell.getNumericCellValue());
							break;
						case STRING:
							newCell.setCellValue(oldCell.getRichStringCellValue());
							break;
						}
					}
					if(i<1) {
						// If there are are any merged regions in the source row, copy to new row
						for (int h = 0; h < sheet.getNumMergedRegions(); h++) {
							CellRangeAddress cellRangeAddress = sheet.getMergedRegion(h);
							if (cellRangeAddress.getFirstRow() == sourceRow.getRowNum()) {
								CellRangeAddress newCellRangeAddress = new CellRangeAddress(newRow.getRowNum(),
										(newRow.getRowNum() + (cellRangeAddress.getLastRow() - cellRangeAddress.getFirstRow())),
										cellRangeAddress.getFirstColumn(), cellRangeAddress.getLastColumn());
								sheetTemp.addMergedRegion(newCellRangeAddress);
							}
						}
					}
					contadorDestino=contadorDestino+3;
					contador++;
				}
			}
			//Fim Corpo XML para Planilha
			//FIM AJUSTAR LAYOUT REPORT
			
			workbookReport.removeSheetAt(0);
			sheetTemp.shiftRows(0, sheetTemp.getLastRowNum(), 1, true, false);

			out = new FileOutputStream(fileReport);

			workbookReport.write(out);
			out.close();
			workbookReport.close();
			//FIM AJUSTAR LAYOUT REPORT
			*/
			//INICIO JUNTAR REPORTS
			
			//To DELETE
			ArrayList<XSSFWorkbook> workbooks = new ArrayList<XSSFWorkbook>();
			
			File xmlFile = new File(caminhoPastaXML + nome_planilha_xml);
			System.out.println(xmlFile.getAbsolutePath());
			ZipSecureFile.setMinInflateRatio(0);
			workbooks = null;
			workbooks.add(new XSSFWorkbook(new FileInputStream(xmlFile.getAbsolutePath())));
			
			File reportFile = new File(path + "/report/" + nome_report);
			System.out.println(reportFile.getAbsolutePath());
			workbooks.add(new XSSFWorkbook(new FileInputStream(reportFile.getAbsolutePath())));
			
			File templateFile = new File(caminhoPastaXML+"template_novo.xlsx");
			XSSFWorkbook workbookTemplate = new XSSFWorkbook(new FileInputStream(templateFile.getAbsolutePath()));
			XSSFSheet sheetTemplate = workbookTemplate.getSheetAt(0); 
			
			Map<Integer, ArrayList<String>> reportData = new TreeMap<Integer, ArrayList<String>>();
			ArrayList<String> listaPlanilha1;
			ArrayList<String> listaPlanilha2;
			ArrayList<String> listaPlanilha3;
			
			for (int k = 0; k < workbooks.size(); k++) {
				for (int i = 0; i < workbooks.get(k).getSheetAt(0).getPhysicalNumberOfRows(); i++) {
					XSSFRow sourceRow = workbooks.get(k).getSheetAt(0).getRow(i);
					// Loop through source columns to add to new row
					for (int j = 0; j < sourceRow.getLastCellNum(); j++) {
						// Grab a copy of the old/new cell
						XSSFCell oldCell = sourceRow.getCell(j);

						/*
						// Copy style from old cell and apply to new cell
						XSSFCellStyle newCellStyle = workbookReport.createCellStyle();
						newCellStyle.cloneStyleFrom(oldCell.getCellStyle());
						newCell.setCellStyle(newCellStyle);

						// If there is a cell comment, copy
						if (oldCell.getCellComment() != null) {
							newCell.setCellComment(oldCell.getCellComment());
						}

						// If there is a cell hyperlink, copy
						if (oldCell.getHyperlink() != null) {
							newCell.setHyperlink(oldCell.getHyperlink());
						}

						// Set the cell data type
						newCell.setCellType(oldCell.getCellType());*/

						// Set the cell data value
						switch (oldCell.getCellType()) {
						case BLANK:
							tempObject.add(oldCell.getStringCellValue());
							break;
						/*case BOOLEAN:
							tempObject.add(oldCell.getBooleanCellValue());
							break;
						case ERROR:
							tempObject.add(oldCell.getErrorCellValue());
							break;
						case FORMULA:
							tempObject.add(oldCell.getCellFormula());
							break;
						case NUMERIC:
							tempObject.add(oldCell.getNumericCellValue());
							break;*/
						case STRING:
							tempObject.add(oldCell.getRichStringCellValue().toString());
							break;
						}
					}
				}
			}

			
			FileOutputStream out = new FileOutputStream(templateFile);

			workbookTemplate.write(out);

			out.close();
			//workbookReport.close();         DESCOMENTAR NO FINAL
			workbookTemplate.close();

			
					
			System.out.println("Acabou aqui!!!!!!!");
		}

		catch (Exception e) {
			System.out.println("Houve um problema com o aplicativo.");
			System.out.println(e);
			e.printStackTrace();
		}

	}

	public static void uploadFile(String fileLocation) {
		try {
			setClipboardData(fileLocation);
			Robot robot = new Robot();
			Thread.sleep(3000);
			robot.keyPress(KeyEvent.VK_ALT);
			robot.keyRelease(KeyEvent.VK_ALT);
			Thread.sleep(1000);
			robot.keyPress(KeyEvent.VK_CONTROL);
			robot.keyPress(KeyEvent.VK_V);
			robot.keyRelease(KeyEvent.VK_V);
			robot.keyRelease(KeyEvent.VK_CONTROL);
			robot.keyPress(KeyEvent.VK_ENTER);
			Thread.sleep(3000);
			robot.keyPress(KeyEvent.VK_ENTER);
		} catch (Exception exp) {
			exp.printStackTrace();
		}
	}
	
	public static void setClipboardData(String string) {
		StringSelection stringSelection = new StringSelection(string);
		Toolkit.getDefaultToolkit().getSystemClipboard().setContents(stringSelection, null);
	}

	public static void saveBackupFile(String report) {
		String nameFile = path + name + ".xlsb";

		try {
			String[] command = new String[5];
			command[0] = "cmd";
			command[1] = "/c";
			command[2] = "copy";
			command[3] = nameFile;
			command[4] = report;
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

	public static void deleteFilesXLSB() {
		String nameFile = path + name + ".xlsb";
		File f = new File(nameFile);
		f.delete();
	}
	private static boolean isRowEmpty(XSSFRow row) {
		boolean isEmpty = true;
		DataFormatter dataFormatter = new DataFormatter();

		if (row != null) {
			for (Cell cell : row) {
				if (dataFormatter.formatCellValue(cell).trim().length() > 0) {
					isEmpty = false;
					break;
				}
			}
		}

		return isEmpty;
	}

}