package factoryDesignPattern;

import java.io.File;
import java.io.IOException;
import java.util.Scanner;

import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;

public class readExceluserInput {

	public static void main(String[] args) throws IOException {
		Scanner sc = new Scanner(System.in);
		System.out.println("Enter row number");
		int rowIndex = sc.nextInt();
		readExcel(rowIndex);
	}

	private static void readExcel(int rowIndex) throws IOException {
		File file = new File("C:/Users/jemalmohammed/Documents/Test.xls");
        Workbook w;
        try {
                w = Workbook.getWorkbook(file);
                // Get the first sheet
                Sheet sheet = w.getSheet(0);
                
                //first 5 from top
                String result ="";
                for (int col = 0; col < sheet.getColumns(); col++) {
                     Cell cell = sheet.getCell(col, rowIndex);
                     result += cell.getContents() +" ";
                }
                System.out.println( result);
                
                
        } catch (BiffException e) {
                e.printStackTrace();
        }
		
	}

}
