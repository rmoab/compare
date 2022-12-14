package br.com.validation;

import java.io.File;
import java.io.FileInputStream;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Iterator;
import java.util.Map;
import java.util.TreeMap;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadExcelDemo 
{
	
	static String pathFile;
	static String workSheetName;
	static String pathFileEnd;
	static Object[] obj = new Object[] {};
	static ArrayList<Object> newObj = new ArrayList<Object>(Arrays.asList(obj)); 
	
	   public static void main(String[] args) throws Exception {
		   
		   for (int i = 0; i < args.length; i++) {
				System.out.println(args[i]);
			}
		   
		   pathFile = (args[0]);
		   workSheetName = (args[1]);
		   pathFileEnd = (args[2]);
		   		   
//	        String excelFilePath = "C:/Temp/RETEST_NEW_ SAZ_Brazil_CANS 0203_Large RFP 1.xlsx";
		    String excelFilePath = pathFile;	   
	        FileInputStream inputStream = new FileInputStream(new File(excelFilePath));
	         
	        Workbook workbook = new XSSFWorkbook(inputStream);
//	        Sheet firstSheet = workbook.getSheetAt(1);
	        Sheet firstSheet = workbook.getSheet(workSheetName);
	        Iterator<Row> iterator = firstSheet.iterator();
	        
	        DataFormatter df = new DataFormatter();

	        df.addFormat("General", new java.text.DecimalFormat("#.###############"));
	               
	        String celula = "";
	        boolean celulaBoolean = false;
	        double celulaNumeric = 0;
	        int numeracao = 1;
	        
	        Map<Integer, Object[]> reportData = new TreeMap<Integer, Object[]>();	        
	       
	        while (iterator.hasNext()) {
	            Row nextRow = iterator.next();
	            Iterator<Cell> cellIterator = nextRow.cellIterator();
	            newObj.clear();

	            int rowNumber = nextRow.getRowNum();
	            
	            if(rowNumber!=0) {
	            
	            while (cellIterator.hasNext()) {
	                Cell cell = cellIterator.next();
	               
		           	                
	                switch (cell.getCellType()) {
	                    case STRING:
	                    	celula = cell.getStringCellValue();
	                    	newObj.add(celula);
	                        break;
	                    case BOOLEAN:
	                    	celulaBoolean = cell.getBooleanCellValue();
	                    	newObj.add(celulaBoolean);
	                        break;
	                    case NUMERIC:            	
	                    	celulaNumeric = cell.getNumericCellValue();                    	
	                    	String value = df.formatCellValue(cell);
	                    	newObj.add(value);
	                        break;
	                    default:
	                    	newObj.add(cell);
	                }

	                
	            }
	             
	            String valor1 = trataValor(0);
	            String valor2 = trataValor(1);
	            String valor3 = trataValor(2);
	            String valor4 = trataValor(3);
	            String valor5 = trataValor(4);
	            String valor6 = trataValor(5);
	            String valor7 = trataValor(6);
	            String valor8 = trataValor(7);
	            String valor9 = trataValor(8);
	            String valor10;
	            
	            if(workSheetName.equals("Item Attributes")) {
	            	valor10 = "";
	            }else {
	            	valor10 = trataValor(9);
	            }
	            
	            String valor11 = trataValor(10);
	            String valor12 = trataValor(11);
	            String valor13 = trataValor(12);
	            String valor14 = trataValor(13);
                String valor15 = trataValor(14);
	            String valor16 = trataValor(15);
	            String valor17 = trataValor(16);
	            String valor18 = trataValor(17);
	            String valor19 = trataValor(18);
	            String valor20 = trataValor(19);
	            String valor21 = trataValor(20);
	            String valor22 = trataValor(21);
	            String valor23 = trataValor(22);
	            String valor24 = trataValor(23);
	            String valor25 = trataValor(24);
	            String valor26 = trataValor(25);
	            String valor27 = trataValor(26);
	            String valor28 = trataValor(27);
	             
	            reportData.put(numeracao, new Object[]{valor1, valor2, valor3, valor4, valor5, valor6, valor7, valor8, valor9, valor10,
	            		valor11, valor12, valor13, valor14, valor15, valor16, valor17, valor18, valor19, valor20,	
	            		valor21, valor22, valor23, valor24, valor25, valor26, valor27, valor28
	            });

	            numeracao = numeracao + 1;      	
	    
	        }
	        } 
	        
	           UpdateExistingExcelFile updateExistingExcelFile = new UpdateExistingExcelFile();
	           UpdateExistingExcelFile.update(reportData, pathFileEnd, workSheetName); 
	            
	        workbook.close();
	        inputStream.close();
	    }

	private static String trataValor(int i) {
		 String valorTratado = "";

		 try {
			 valorTratado = newObj.get(i).toString();		 
		 } catch (IndexOutOfBoundsException e) {
			valorTratado = "";
			}
			 
		return valorTratado;
	}
}
