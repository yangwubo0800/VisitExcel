package main;

import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;

import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;

public class readExcel {

	/**
	 * @param args
	 */
	public static void main(String[] args) {
		// TODO Auto-generated method stub
		
		try {
			File file = new File("f:\\test.xls");  
			InputStream in = new FileInputStream(file);  
			Workbook workbook = Workbook.getWorkbook(in);  
			//获取第一张Sheet表  
			Sheet sheet = workbook.getSheet(0);  
			
			//
			for(int i=0;i<sheet.getRows(); i++){
				for(int j=0; j<sheet.getColumns(); j++){
					Cell cell = sheet.getCell(j, i);
					
					System.out.print(cell.getContents()+" ");
					System.out.println();
					
				}
				
			}
			
			workbook.close();
		}catch(Exception e){
			e.printStackTrace();
			
		}

	}

}
