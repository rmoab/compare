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
import java.util.Scanner;

import javax.validation.constraints.AssertTrue;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.util.IOUtils;


public class OpenExcelValidate {

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
			
			
			/*
			ArrayList<String> nomesArquivosXML = new ArrayList<String>();
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
					
					// reexibir planilha LerXml
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
					Thread.sleep(1000);
					robot.keyPress(KeyEvent.VK_DOWN);
					robot.keyRelease(KeyEvent.VK_DOWN);
					robot.keyPress(KeyEvent.VK_DOWN);
					robot.keyRelease(KeyEvent.VK_DOWN);
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
					Thread.sleep(130000);

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

			*/
			
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
			System.out.println("Workbook: "+workbook);
			System.out.println("Diretorio: "+diretorio);
			ArrayList<XSSFSheet> sheets = new ArrayList<XSSFSheet>();
			ArrayList<Integer> rows = new ArrayList<Integer>(); 
			int contadorXML = 0;
			int contadorDestino = 0;
			int contadorSource = 0;
			int h=0;
			for (File file : diretorio.listFiles()) {	
				System.out.println("file name: "+file.getName());
				if (file.getName().contains("lerXML_")) {
					System.out.println("IF contains lerXML");
					System.out.println(file.getAbsolutePath());
				    XSSFWorkbook tempWorkbook = new XSSFWorkbook(new FileInputStream(file.getAbsolutePath()));
				    XSSFSheet tempSheet = tempWorkbook.getSheet("LerXML");
				    contadorSource = tempSheet.getLastRowNum();
				    for(int contadorLinha = 0;contadorLinha<contadorSource;contadorLinha++) {
						// Get the source / new row
						XSSFRow sourceRow = sheets.get(contadorXML).getRow(contadorLinha);
						XSSFRow newRow = destinationWorksheet.createRow(contadorDestino);
						// Loop through source columns to add to new row
						h=0;
						for (int i = 0; i < sourceRow.getLastCellNum(); i++) {
							if(contadorXML>0) {
								i=2;
								h=2+1;
							}
							// Grab a copy of the old/new cell
							XSSFCell oldCell = sourceRow.getCell(i);
							XSSFCell newCell = newRow.createCell(h);
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
					        h++;
						}

						// If there are are any merged regions in the source row, copy to new row
						for (int i = 0; i < sheets.get(contadorXML).getNumMergedRegions(); i++) {
							CellRangeAddress cellRangeAddress = sheets.get(contadorXML).getMergedRegion(i);
							if (cellRangeAddress.getFirstRow() == sourceRow.getRowNum()) {
								CellRangeAddress newCellRangeAddress = new CellRangeAddress(newRow.getRowNum(),
										(newRow.getRowNum() + (cellRangeAddress.getLastRow() - cellRangeAddress.getFirstRow())),
										cellRangeAddress.getFirstColumn(), cellRangeAddress.getLastColumn());
								destinationWorksheet.addMergedRegion(newCellRangeAddress);
							}
						}
				    tempWorkbook.close();
				    System.out.println("Fim IF contains lerXML");
				    contadorXML++;
				}
				
			}
			
			/*Integer quantidadeXMLs = sheets.size();
			int contadorDestino = 0;		
			for (int contadorXML = 0; contadorXML < quantidadeXMLs; contadorXML++) {
				//Para o primeiro caso pegar o cabeçalho linha 1 e 2. Para os demais pegar a partir da linha 2;
				for (int contadorLinhas = 0; contadorLinhas < rows.get(contadorXML); contadorLinhas++) {
					if(contadorXML>0) {
						contadorLinhas=2;
						contadorDestino=contadorLinhas+1;
					}
					// Get the source / new row
					XSSFRow sourceRow = sheets.get(contadorXML).getRow(contadorLinhas);
					XSSFRow newRow = destinationWorksheet.createRow(contadorDestino);
					// Loop through source columns to add to new row
					for (int i = 0; i < sourceRow.getLastCellNum(); i++) {
						// Grab a copy of the old/new cell
						XSSFCell oldCell = sourceRow.getCell(i);
						XSSFCell newCell = newRow.createCell(i);
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
					for (int i = 0; i < sheets.get(contadorXML).getNumMergedRegions(); i++) {
						CellRangeAddress cellRangeAddress = sheets.get(contadorXML).getMergedRegion(i);
						if (cellRangeAddress.getFirstRow() == sourceRow.getRowNum()) {
							CellRangeAddress newCellRangeAddress = new CellRangeAddress(newRow.getRowNum(),
									(newRow.getRowNum() + (cellRangeAddress.getLastRow() - cellRangeAddress.getFirstRow())),
									cellRangeAddress.getFirstColumn(), cellRangeAddress.getLastColumn());
							destinationWorksheet.addMergedRegion(newCellRangeAddress);
						}
					}
				}
			}*/
			// .xlsx is the format for Excel Sheets...
			// writing the workbook into the file...
			FileOutputStream out = new FileOutputStream(new File(path + "/report/xml.xlsx"));

			workbook.write(out);

			out.close();
			workbook.close();
			
			System.out.println("Acabou aqui!!!!!!!");
			
			
			
			
			
			
			
			
			
			
			
			
			
			
			
			
			
			
			
			
			
			
			
			
			
			
			
			
			
			
			
			
			
			
			
			
			
			
			/*
			// Localiza todas os XMLS da pasta
			for (File file1 : diretorio.listFiles()) {
				if (file1.getName().contains(".xml")) {
					nomeArquivos.add(file1.getAbsolutePath());
					name = file1.getName();
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
					uploadFile(file1.getAbsolutePath());
					Thread.sleep(20000);
					robot.keyPress(KeyEvent.VK_ALT);
					Thread.sleep(2500);
					robot.keyPress(KeyEvent.VK_S);
					
					Thread.sleep(1000);
					robot.keyPress(KeyEvent.VK_ALT);
					robot.keyRelease(KeyEvent.VK_ALT);
					Thread.sleep(3000);

					// Salvar arquivo validado
					robot.keyPress(KeyEvent.VK_F12);
					robot.keyRelease(KeyEvent.VK_F12);
					Thread.sleep(3000);
					// criar pasta
					setClipboardData(name);
					Thread.sleep(1000);
					robot.keyPress(KeyEvent.VK_CONTROL);
					robot.keyPress(KeyEvent.VK_V);
					robot.keyRelease(KeyEvent.VK_V);
					robot.keyRelease(KeyEvent.VK_CONTROL);
					robot.keyPress(KeyEvent.VK_ENTER);
					Thread.sleep(3000);
					robot.keyPress(KeyEvent.VK_ENTER);
					Thread.sleep(130000);

					// exportar excel em TXT
					// Salvar arquivo validado em TXT
					robot.keyPress(KeyEvent.VK_F12);
					robot.keyRelease(KeyEvent.VK_F12);
					Thread.sleep(3000);
					robot.keyPress(KeyEvent.VK_TAB);
					robot.keyRelease(KeyEvent.VK_TAB);
					robot.keyPress(KeyEvent.VK_T);
					robot.keyRelease(KeyEvent.VK_T);
					Thread.sleep(3000);
					robot.keyPress(KeyEvent.VK_ENTER);
					Thread.sleep(3000);
					robot.keyPress(KeyEvent.VK_ENTER);
					Thread.sleep(30000);
					
					// reexibir planilha LerXml
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
					Thread.sleep(1000);
					robot.keyPress(KeyEvent.VK_DOWN);
					robot.keyRelease(KeyEvent.VK_DOWN);
					robot.keyPress(KeyEvent.VK_DOWN);
					robot.keyRelease(KeyEvent.VK_DOWN);
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
					
					// Salvar report
					robot.keyPress(KeyEvent.VK_F12);
					robot.keyRelease(KeyEvent.VK_F12);
					Thread.sleep(3000);
					// criar pasta
					setClipboardData(reportPath+name);
					Thread.sleep(1000);
					robot.keyPress(KeyEvent.VK_CONTROL);
					robot.keyPress(KeyEvent.VK_V);
					robot.keyRelease(KeyEvent.VK_V);
					robot.keyRelease(KeyEvent.VK_CONTROL);
					//selecionar tipo do arquivo
					robot.keyPress(KeyEvent.VK_TAB);
					robot.keyRelease(KeyEvent.VK_TAB);
					robot.keyPress(KeyEvent.VK_LEFT);
					robot.keyRelease(KeyEvent.VK_LEFT);
					System.out.println("Selecionar tipo Arquivo");
					for(int i = 1;i<=9;i++) {
						Thread.sleep(1000);
					robot.keyPress(KeyEvent.VK_UP);
					robot.keyRelease(KeyEvent.VK_UP);
					}
					//salvar
					robot.keyPress(KeyEvent.VK_ENTER);
					Thread.sleep(3000);
					robot.keyPress(KeyEvent.VK_ENTER);
					Thread.sleep(130000);
					

					// Fechar arquivo
					robot.keyPress(KeyEvent.VK_ALT);
					robot.keyPress(KeyEvent.VK_F4);
					Thread.sleep(1500);
					Thread.sleep(3000);
					robot.keyPress(KeyEvent.VK_ENTER);
					Thread.sleep(1500);
					robot.keyPress(KeyEvent.VK_ALT);
					robot.keyRelease(KeyEvent.VK_ALT);
					robot.keyPress(KeyEvent.VK_N);
					robot.keyPress(KeyEvent.VK_N);
					Thread.sleep(5000);

					// Valida se teve divergência
					ValidateTXT validateTXT = new ValidateTXT();
					boolean retorno = validateTXT.validarTXT(name, caminhoPastaXML, arquivoValidador);

//			    System.out.println("Retorno de divergencia é " + retorno);

					if (retorno == true) {
						listaArquivosDivergencia.add(name);

						// Salvar arquivo com divergência na pasta de divergências
						saveBackupFile(reportPathDivergencia);
					} else {
						saveBackupFile(reportPath);
					}

					deleteFilesXLSB();
//			    
				}

			}*/
			System.out.println("Validação dos arquivos concluida com sucesso.");

			// chama ExcelData
			ExcelData excelData = new ExcelData();
			excelData.writeExcel(data, listaArquivosDivergencia, arquivoValidador, caminhoPastaXML);

//	    	deleteFilesTXT();

			if (listaArquivosDivergencia.size() > 0) {
				System.out.println("Os arquivos com divergência são: \r\n");
				for (int i = 0; i < listaArquivosDivergencia.size(); i++) {
					if (listaArquivosDivergencia.get(i).equals(listaArquivosDivergencia.get(i))) {
						System.out.println(listaArquivosDivergencia.get(i));
					}
				}
			} else {
//				System.out.println("Não existem arquivos com divergência");
			}
		}
			}

		catch (Exception e) {
			System.out.println("Houve um problema com o aplicativo.");
			System.out.println(e);
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

}