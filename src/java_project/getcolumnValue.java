package factoryDesignPattern;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

import jxl.Cell;
import jxl.CellType;
import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;

public class Fercode {

	public static void main(String[] args) throws IOException {
		readExcelFile();

	}

	
	public static void readExcelFile() throws IOException {
		File file = new File("C:/Users/jemalmohammed/Documents/Test.xls");
        Workbook w;
        try {
                w = Workbook.getWorkbook(file);
                // Get the first sheet
                Sheet sheet = w.getSheet(0);
                
                //first 5 from top
                String result ="";
                for (int i = 0; i < 5; i++) {
                    for (int j = 0; j< sheet.getColumns(); j++) {
                            Cell cell = sheet.getCell(j, i);
                             result += cell.getContents() +" ";
                            }

                    result+="\n";
                    }
                System.out.println( result);
                ////////////////////////////////
                
              //first 5 from bottom
                result ="";
                int numOfRows = sheet.getRows();
                for (int i = sheet.getRows()-1; i>numOfRows-6 ; i--) {
                    for (int j = 0; j< sheet.getColumns(); j++) {
                            Cell cell = sheet.getCell(j, i);
                            
                                    //System.out.println( "Row" + j);
                                    result += cell.getContents() +" ";
                            }

                    result+="\n";
                    }
                System.out.println( result);
                ////////////////////////
                
                
        } catch (BiffException e) {
                e.printStackTrace();
        }
	}

}
