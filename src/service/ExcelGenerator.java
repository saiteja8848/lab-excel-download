package service;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import model.Prograd;

//			Progression -1 
//Go to src/service. Open the ExcelGenerator and fill the logic inside the excelGenerate method.
//
//Stick to the instructions clearly. If you face any issue contact your mentor to get the guidance. 

public class ExcelGenerator {
	
	FileOutputStream out;

	
	public HSSFWorkbook excelGenerate(Prograd prograd, List<Prograd> list) throws IOException {
		
		HSSFWorkbook workbook=null;
		FileOutputStream out=null;
		HSSFSheet worksheet;
		HSSFRow rows; 
		try {
			String filename="prograd.xls";
			
			workbook = new HSSFWorkbook();	
			worksheet =workbook.createSheet();
			
			for (int r=0;r < list.size(); r++ )
			{
				HSSFRow row = worksheet.createRow(r);
				for (int c=0;c < 5; c++ )
				{
					HSSFCell cell = row.createCell(c);		
					
					if(c==0)
					cell.setCellValue(list.get(r).getName());
					else if(c==1)
					cell.setCellValue(list.get(r).getId());
					else if(c==2)
					cell.setCellValue(list.get(r).getRate());
					else if(c==3)
					cell.setCellValue(list.get(r).getComment());
					else
						if(c==4)
					cell.setCellValue(list.get(r).getRecommend());
				}
			}
			
			out = new FileOutputStream(filename);
			workbook.write(out);	
		}
			catch (Exception e) {e.printStackTrace();}
		finally {
			out.close();}
		
			 
		return workbook;		
	}

}
