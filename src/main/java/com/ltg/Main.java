package com.ltg;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.TreeMap;
import java.util.stream.Stream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Main {

	static XSSFRow row;

	@SuppressWarnings("unused")
	public static void main(String[] args) {
		try {
			final String readPath = args[0];
			final String writePath = args[1];
			
//			final String readPath = "C:\\Users\\emylyn.audemard\\Downloads\\test\\output";
//			final String writePath = "C:\\Users\\emylyn.audemard\\Downloads\\test";
			
			DateFormat df = new SimpleDateFormat("MM/dd/yyyy");
			List<List<Object>> lists = new ArrayList<List<Object>>();
			String masterFilename = "master";
			
			String companyName[];
		    String LTGCompany;
		    
			if(dirIsEmpty(readPath)) {
				System.out.println("Error....");
				System.out.println("Directory is empty.");
				
				System.exit(0);
			}
			
			File directoryPath = new File(readPath);
			String contents[] = directoryPath.list();
			
			System.out.println("List of files in the specified directory:");
			
			for(int i=0; i<contents.length; i++) {
				
				int record = 0;
				
				File file = new File(directoryPath+"\\"+contents[i]);
				FileInputStream fis = new FileInputStream(file);
				XSSFWorkbook workbook = new XSSFWorkbook(fis);
			    XSSFSheet spreadsheet = workbook.getSheetAt(0);
				
				String filename = file.getName().substring(0, file.getName().lastIndexOf("."));
			
				if(filename.contains("3P")){
					filename = filename+"-3P";
				}
				
				companyName = filename.split("_");
				LTGCompany = companyName[1].toString().trim();
				
				System.out.print(file.getName()+" Processing.......");
				
				Iterator < Row >  rowIterator = spreadsheet.iterator();
				
				while (rowIterator.hasNext()) {
					Row row = rowIterator.next();
					
					Iterator<Cell> cellIterator = row.cellIterator();
					List<Object> list = new ArrayList<Object>(); 
					Map<Integer, Object> map=new HashMap<Integer, Object>();  
					
					while (cellIterator.hasNext()) {
						Cell cell = cellIterator.next();
						
						if(cell.getColumnIndex()==0) { //Category
							switch (cell.getCellType()) {
								case STRING:
									list.add(cell.getStringCellValue());
									map.put(1, cell.getStringCellValue());
									break;
								case BLANK:
									list.add("aaa");
									break;
								default:
									list.add("def");
									break;
							}
						}
						
						if(cell.getColumnIndex()==11) { //Region
							switch (cell.getCellType()) {
								case STRING:
									list.add(cell.getStringCellValue());
									map.put(2, cell.getStringCellValue());
									break;
								case BLANK:
									list.add("aaa");
									break;
								default:
									list.add("def");
									break;
							}
						}
						
						if(cell.getColumnIndex()==12) { //Province
							switch (cell.getCellType()) {
								case STRING:
									list.add(cell.getStringCellValue());
									map.put(3, cell.getStringCellValue());
									break;
								case BLANK:
									list.add("aaa");
									break;
								default:
									list.add("def");
									break;
							}
						}
						
						if(cell.getColumnIndex()==13) { //city
							switch (cell.getCellType()) {
								case STRING:
									list.add(cell.getStringCellValue());
									map.put(4, cell.getStringCellValue());
									break;
								case BLANK:
									list.add("aaa");
									break;
								default:
									list.add("def");
									break;
							}
						}
						
						if(cell.getColumnIndex()==14) { //Bgy
							switch (cell.getCellType()) {
								case STRING:
									list.add(cell.getStringCellValue());
									map.put(5, cell.getStringCellValue());
									break;
								case BLANK:
									list.add("aaa");
									break;
								default:
									list.add("def");
									break;
							}
						}
						
						if(cell.getColumnIndex()==21) { //employer
							switch (cell.getCellType()) {
								case STRING:
									list.add(cell.getStringCellValue());
									map.put(6, cell.getStringCellValue());
									break;
								case BLANK:
									list.add("aaa");
									break;
								default:
									list.add("def");
									break;
							}
						}
						
						if(cell.getColumnIndex()==45) { //Willing
							switch (cell.getCellType()) {
								case STRING:
									list.add(cell.getStringCellValue());
									map.put(7, cell.getStringCellValue());
									break;
								case BLANK:
									list.add("aaa");
									break;
								default:
									list.add("def");
									break;
							}
						}
						
						if(cell.getColumnIndex()==53) { //completion time
							switch (cell.getCellType()) {
								case STRING:
									list.add(cell.getStringCellValue());
									map.put(8, cell.getStringCellValue());
									break;
								case NUMERIC:
									list.add(df.format(cell.getDateCellValue()));
									map.put(8, cell.getDateCellValue());
									break;
								case BLANK:
									list.add("aaa");
									break;
								default:
									list.add("def");
									break;
							}
						}
						
						if(cell.getColumnIndex()==61) { //Type of Employee - to determine 3p
							switch (cell.getCellType()) {
								case STRING:
									list.add(cell.getStringCellValue());
									map.put(9, cell.getStringCellValue());
									break;
								case BLANK:
									list.add("aaa");
									break;
								default:
									list.add("def");
									break;
							}
						}
						
					}
					
					lists.add(list);
					record++;
//					for (List<Object> mylist : lists) {  
//						System.out.println(mylist);
//				    } 
				}
				
				lists.removeIf(e -> e.contains("Category"));//Filter: remove headers contains category keyword
				lists.removeIf(e -> e.contains("Category*"));//Filter: remove headers contains category keyword
				fis.close();
			    workbook.close();
			    
			    System.out.println("DONE "+record);
				record=0;
				
				
			}
			System.out.println();
		    System.out.println("Total Lists:"+lists.size());
		    
		    Map < String, List<Object> > reportinfo = new TreeMap < String, List<Object> >();
		    
		    int item = 1;
		    for (List<Object> list : lists) {  
		    	if(!list.isEmpty()) {
		    		reportinfo.put(String.valueOf(item), list);
		    		item++;
		    	}
		    } 
		    
		    XSSFWorkbook workbook_out = new XSSFWorkbook(); 
		    XSSFSheet spreadsheet_out = workbook_out.createSheet(" Consolidated Result ");
		    
		  //Iterate over data and write to sheet
		      Set < String > keyid = reportinfo.keySet();
		      int rowid = 0;
		      
		      for (String key : keyid) {
			         row = spreadsheet_out.createRow(rowid++);
			         List<Object> objectArr = reportinfo.get(key);
			         int cellid = 0;

			         for (Object obj : objectArr) {
			            Cell cell = row.createCell(cellid++);
			            
			            if(obj.toString().isEmpty()){
			            	cell.setCellValue("--");
			            }else if(obj.toString().trim() == "aaa"){
			            	cell.setCellValue("--");
			            }else {
			            	cell.setCellValue(obj.toString());
			            }
			            
			         }
			      }
		      
			    //Write the workbook in file system
		      FileOutputStream out = new FileOutputStream(new File(writePath+"\\"+masterFilename+".xlsx"));
//		      FileOutputStream out = new FileOutputStream(new File(writePath+"\\"+masterFilename+"_"+LocalDateTime.now().format(DateTimeFormatter.ofPattern("ddMMyyyy_HHmmss"))+".xlsx"));
		      workbook_out.write(out);
		      out.close();
		      workbook_out.close();
		      System.out.println("written successfully");
			
	    }
	    catch (ArrayIndexOutOfBoundsException | IOException e){
	        System.out.println(e);
	    }
	    finally {

	    }
	}
	
	static Boolean checkFolder(String path) {
		File f = new File(path);
		
		return f.exists() && f.isDirectory();
		
	}
	
	public static boolean dirIsEmpty(String path) throws IOException {
		Path p = Paths.get(path);
		
	    if (Files.isDirectory(p)) {
	        try (Stream<Path> entries = Files.list(p)) {
	            return !entries.findFirst().isPresent();
	        }
	    }
	        
	    return false;
	}

}
