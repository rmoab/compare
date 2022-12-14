package br.com.validation;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Map;
import java.util.Set;
import java.util.TreeMap;
import java.util.TreeSet;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
 
public class UpdateExistingExcelFile {
 
//    public static void main(String[] args) {
    	
    	public static void update(Map<Integer, Object[]> newObj, String pathFileEnd, String workSheetName) throws Exception {
 
        // Creating file object of existing excel file
//      File xlsxFile = new File("C:/Temp/Line Items Document.xlsx");    
        File xlsxFile = new File(pathFileEnd); 
     
        Map<Integer, Object[]> reportData = new TreeMap<Integer, Object[]>();
 		reportData = newObj;

        try {
            //Creating input stream
            FileInputStream inputStreamUpdate = new FileInputStream(xlsxFile);
             
            //Creating workbook from input stream
            Workbook workbookUpdate = new XSSFWorkbook(inputStreamUpdate);
//            Workbook workbookUpdate = WorkbookFactory.create(inputStreamUpdate);
 
            //Reading first sheet of excel file
//            Sheet sheet = workbook.getSheetAt(1);
            Sheet sheet = workbookUpdate.getSheet(workSheetName);
 
            //Getting the count of existing records
            int rowCount = sheet.getLastRowNum();
 
            //Iterating new students to update
//            for (Object[] student : newStudents) {
            
            Set<Integer> keyid = new TreeSet<Integer>();
    		keyid = newObj.keySet();

    		int rowid = 0;

//            int rowid = count;

    		// writing the data into the sheets...

    		for (Integer key : keyid) {
    			
    			Row row = sheet.createRow(++rowCount);
    			Object[] objectArr = reportData.get(key);
    			int cellid = 0;

    			for (Object obj : objectArr) {
    				Cell cell = row.createCell(cellid++);
    				cell.setCellValue((String) obj);


    			}
    		}

            //Close input stream
    		inputStreamUpdate.close();
 
            //Crating output stream and writing the updated workbook
            FileOutputStream os = new FileOutputStream(xlsxFile);
            workbookUpdate.write(os);
             
            //Close the workbook and output stream
            workbookUpdate.close();
            os.close();
             
            System.out.println("Planilha do Excel foi atualizada com sucesso.");
             
        } catch (EncryptedDocumentException | IOException e) {
            System.err.println("Ocorreu um erro ao atualizar a Planilha do Excel");
            e.printStackTrace();
        }
    }

		
}