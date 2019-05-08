package com.fusu;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;

public class TestExcel {
	@Test
	public void testAdd() throws IOException {
		/**
		 *  HSSF － 提供读写Microsoft Excel XLS格式档案的功能。
			XSSF － 提供读写Microsoft Excel OOXML XLSX格式档案的功能。
			HWPF － 提供读写Microsoft Word DOC97格式档案的功能。
			XWPF － 提供读写Microsoft Word DOC2003格式档案的功能。
			HSLF － 提供读写Microsoft PowerPoint格式档案的功能。
			HDGF － 提供读Microsoft Visio格式档案的功能。
			HPBF － 提供读Microsoft Publisher格式档案的功能。
			HSMF － 提供读Microsoft Outlook格式档案的功能。
		 */
		// 第一步创建workbook
//		HSSFWorkbook wb = new HSSFWorkbook();
		XSSFWorkbook wb = new XSSFWorkbook();
		// 第二步创建sheet
		XSSFSheet sheet = wb.createSheet("测试");
//		HSSFSheet sheet = wb.createSheet("测试");

		
		// 第三步创建行row:添加表头0行
//		HSSFRow row = sheet.createRow(0);
		XSSFRow row = sheet.createRow(0);
//		HSSFCellStyle style = wb.createCellStyle();
		XSSFCellStyle style = wb.createCellStyle();
//		style.setAlignment(HSSFCellStyle.ALIGN_CENTER); //居中
		style.setAlignment(XSSFCellStyle.ALIGN_CENTER);
		// 第四步创建单元格
//		HSSFCell cell = row.createCell(0); // 第一个单元格
		XSSFCell cell = row.createCell(0);
		cell.setCellValue("姓名");
		cell.setCellStyle(style);

		cell = row.createCell(1); // 第二个单元格
		cell.setCellValue("年龄");
		cell.setCellStyle(style);

		// 第五步插入数据

		for (int i = 0; i < 10; i++) {
			// 创建行
			row = sheet.createRow(i + 1);
			// 创建单元格并且添加数据
			row.createCell(0).setCellValue("aa" + i);
			row.createCell(1).setCellValue(i);

		}

		// 第六步将生成excel文件保存到指定路径下
		try {
			FileOutputStream fout = new FileOutputStream("D:\\Test\\a.xls");
			wb.write(fout);
			fout.close();
		} catch (IOException e) {
			e.printStackTrace();
		}

		System.out.println("Excel文件生成成功...");
	}
	
	@Test
	public void TestRead() throws Exception {
		String fileName="D:\\Test\\a.xlsx";
        FileInputStream fis = new FileInputStream(fileName); 
          Workbook workbook = null;
        //判断excel的两种格式xls,xlsx
        if(fileName.toLowerCase().endsWith("xlsx")){  
            workbook = new XSSFWorkbook(fis);  
        }else if(fileName.toLowerCase().endsWith("xls")){  
            workbook = new HSSFWorkbook(fis);  
        }  
        
          
        //得到sheet的总数  
        int numberOfSheets = workbook.getNumberOfSheets();  
        
        System.out.println("一共"+numberOfSheets+"个sheet");
        
      //循环每一个sheet  
        for(int i=0; i < numberOfSheets; i++){  
               
            //得到第i个sheet  
            Sheet sheet = workbook.getSheetAt(i);  
            System.out.println(sheet.getSheetName()+"  sheet");
               
            //得到行的迭代器  
            Iterator<Row> rowIterator = sheet.iterator();  
            
            int rowCount=0;
            //循环每一行
            while (rowIterator.hasNext())   
            {  
                System.out.print("第"+(rowCount++)+"行  ");
                
                //得到一行对象  
                Row row = rowIterator.next();  
                   
                //得到列对象 
                Iterator<Cell> cellIterator = row.cellIterator();  
                
                int columnCount=0;  
                
                //循环每一列
                while (cellIterator.hasNext())   
                {  
                    //System.out.print("第"+(columnCount++)+"列:  ");
                    
                    //得到单元格对象
                    Cell cell = cellIterator.next();
                    
                    //检查数据类型 
                    switch(cell.getCellType()){  
                    case Cell.CELL_TYPE_STRING:  
                            System.out.print(cell.getStringCellValue()+"   ");    
                        break;  
                    case Cell.CELL_TYPE_NUMERIC:  
                        System.out.print(cell.getNumericCellValue()+"   ");  
                    }  
                } //end of cell iterator 
                
                System.out.println();
               
            } //end of rows iterator  
            
          
               
        } //end of sheets for loop  
           
        
       System.out.println("\nread excel successfully...");
       
        //close file input stream  
        fis.close();  
	}
}
