package java_project;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.Iterator;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

public class XlstoCSV1 {
	
	

    static void xls(File inputFile, File outputFile) {
        // For storing data into CSV files
        StringBuffer data = new StringBuffer();

       try {
            FileOutputStream fos = new FileOutputStream(outputFile);
            // Get the workbook object for XLS file
            HSSFWorkbook wBook = new HSSFWorkbook(new FileInputStream(inputFile));
            // Get first sheet from the workbook
            HSSFSheet sheet = wBook.getSheetAt(0);
            Row row;
            Cell cell;
            // Iterate through each rows from first sheet
            Iterator<Row> rowIterator = sheet.iterator();

            while (rowIterator.hasNext()) {
                row = rowIterator.next();

                // For each row, iterate through each columns
                Iterator<Cell> cellIterator = row.cellIterator();
                while (cellIterator.hasNext()) {

                    cell = cellIterator.next();

                    switch (cell.getCellType()) {
                        case Cell.CELL_TYPE_BOOLEAN:
                            data.append(cell.getBooleanCellValue() + ",");

                            break;
                        case Cell.CELL_TYPE_NUMERIC:
                            data.append(cell.getNumericCellValue() + ",");

                            break;
                        case Cell.CELL_TYPE_STRING:
                            data.append(cell.getStringCellValue() + ",");
                            break;

                        case Cell.CELL_TYPE_BLANK:
                            data.append("" + ",");
                            break;
                        default:
                            data.append(cell + ",");
                            
                            

                    }
                    
                }
                data.append('\n'); 
            }

            fos.write(data.toString().getBytes());
            fos.close();

        } catch (Exception ioe) {
            ioe.printStackTrace();
        }
    }
    //testing the application 

    public static void main(String[] args) {
    	
    	
        //reading file from desktop
        File inputFile = new File("C:\\Users\\a_ans\\Desktop\\test_data.xls");
        //writing excel data to csv 
        File outputFile = new File("C:\\Users\\a_ans\\Desktop\\output.cvs");
        xls(inputFile, outputFile);
        }

} 
    
    
  
 
    

	






	