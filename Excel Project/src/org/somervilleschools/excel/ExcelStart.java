package org.somervilleschools.excel;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.HashSet;
import java.util.Iterator;
import java.util.TreeMap;

import javax.swing.JFrame;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Row.MissingCellPolicy;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelStart {
	
	HashSet<MainData> mainSet = new HashSet<MainData>();
	HashSet<SecondData> secondSet = new HashSet<SecondData>();
	
	HashMap<String, MainData> mainMap = new HashMap<String, MainData>();
	HashMap<String, SecondData> secondMap = new HashMap<String, SecondData>();
	
	HashSet<Data> noSerialNumber = new HashSet<Data>();
	
	TreeMap<String, DataMatch> serialMatchMap = new TreeMap<String, DataMatch>();
	HashSet<Data> noSerialMatch = new HashSet<Data>();
	
	public static void main(String[] args){	
		new ExcelStart().begin();		
	}
	
	public void begin(){
		
		long start = System.currentTimeMillis();
		
		
		Workbook main = openWorkbook("C:\\Users\\jvandort\\Documents\\Excel Project\\Asset_Export.xlsx");
		Workbook second = openWorkbook("C:\\Users\\jvandort\\Documents\\Excel Project\\Asset_Inventory.xlsx");
		
		//Stops program from throwing errors when it finds an empty cell
		main.setMissingCellPolicy(MissingCellPolicy.CREATE_NULL_AS_BLANK);
		second.setMissingCellPolicy(MissingCellPolicy.CREATE_NULL_AS_BLANK);
		
		readMain(main);
		readSecond(second);
		
		
		System.out.println("Start Compare");
		compare();
		System.out.println("Finished reading in " + (System.currentTimeMillis() - start) + " ms");
		
		start = System.currentTimeMillis();
		
		writeOutput();
		
		System.out.println("Finished writing in " + (System.currentTimeMillis() - start) + " ms");
		
	}
	
	public void writeOutput(){
		Workbook wb = new XSSFWorkbook();
		
		//Write No Serial Number
		Sheet noSerialSheet = createSheet(wb, "No_Serial_Numbers");
		
		int noSerialCtr = 1;
		for(Data data : noSerialNumber){
			Row row = noSerialSheet.createRow(noSerialCtr);
			
			writeData(data, row);
			
			//Next line
			noSerialCtr++;			
		}
		
		//Data Matches by Serial Number
		Sheet serialMatchSheet = createSheet(wb, "Serial Number Matches");
		int serialMatchCtr = 1;
		for(DataMatch match : serialMatchMap.values()){
			
			for(Data data : match){
				Row row = serialMatchSheet.createRow(serialMatchCtr);
				
				writeData(data, row);
				
				//Next line
				serialMatchCtr++;
			}
			
			//Next line
			serialMatchCtr++;			
		}	
		
		//Write no matches
		Sheet noSerialMatchSheet = createSheet(wb, "No Serial Matches");
		int noSerialMatchCtr = 1;
		for(Data data : noSerialMatch){
			Row row = noSerialMatchSheet.createRow(noSerialMatchCtr);
				
			writeData(data, row);
			
			//Next line
			noSerialMatchCtr++;			
		}		
		
		//Adjust Column Widths
		for(Sheet sheet : wb){
			for(int col = 0; col < outputHeaders.length; col++){
				sheet.autoSizeColumn(col);
			}
		}
		
		//Write it
		try {			 
			FileOutputStream fileOut = new FileOutputStream("C:\\Users\\jvandort\\Documents\\Excel Project\\output.xlsx");
			wb.write(fileOut);
			fileOut.close();
			wb.close();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}		
	}
	
	String[] outputHeaders = new String[]{
			"Source Row Num.",
			"Item",
			"Serial Number",
			"Device Name",
			"Asset Tag One",
			"Asset Tag Two",
			"Green Asset Tag",
			"Location",
			"Data Source",
			"Sheet Name"
			};
	
	public Sheet createSheet(Workbook wb, String name){
		Sheet sheet = wb.createSheet(name);		
		writeHeaders(sheet);	
		return sheet;
	}
	
	public void writeHeaders(Sheet sheet){
		sheet.createFreezePane(0, 1, 0, 1);
		Row headerRow = sheet.createRow(0);		 
		
		for(int ctr = 0; ctr < outputHeaders.length; ctr++){
			headerRow.createCell(ctr).setCellValue(outputHeaders[ctr]);
		}
	}
	
	public void writeData(Data data, Row row){
		row.createCell(0).setCellValue(data.getOriginalRowNumber()+1);
		row.createCell(1).setCellValue(data.getItem());
		row.createCell(2).setCellValue(data.getSerialNumber());		
		row.createCell(3).setCellValue(data.getDeviceName());		
		row.createCell(4).setCellValue(data.getAssetTagOne());
		row.createCell(5).setCellValue(data.getAssetTagTwo());
		row.createCell(6).setCellValue(data.getGreenAssetTag());
		row.createCell(7).setCellValue(data.getLocation());
		row.createCell(8).setCellValue(data.getDataSource());
		row.createCell(9).setCellValue(data.getSheetName());		
	}
	
	public void compare(){
		
		for(Data data : mainSet){
			
			String serial = data.getSerialNumber();
			
			if(serialMatchMap.containsKey(serial)){
				serialMatchMap.get(serial).addData(data);
			}else{
				serialMatchMap.put(serial, new DataMatch(data));
			}			
		}
		
		for(Data data : secondSet){
			
			String serial = data.getSerialNumber();
			
			//If there's no serial number, add it to the noSerialNumber set
			if(serial.equals("") || serial.equals("N/A")){
				noSerialNumber.add(data);
				continue;
			}
			
			if(serialMatchMap.containsKey(serial)){
				serialMatchMap.get(serial).addData(data);
			}else{
				serialMatchMap.put(serial, new DataMatch(data));
			}
		}		
		
		//Find and remove Data with no matches for serial Numbers
		//Add move that data to noSerialMatch 
		ArrayList<String> removeList = new ArrayList<String>();		
		for(String serial : serialMatchMap.keySet()){
			
			DataMatch match = serialMatchMap.get(serial);
			
			if(match.size() < 2){
				noSerialMatch.add(match.get(0));
				removeList.add(serial);
			}
		}
		
		for(String remove : removeList){
			serialMatchMap.remove(remove);
		}
	}
	
	public void readMain(Workbook main){
		boolean first = true;
		
		for(Row row : main.getSheetAt(0)){
			
			//Ignore the first row since it has the headers
			if(first){
				first = false;
				continue;
			}
			
			MainData data = new MainData(row);
			
			//After the first row...
			mainSet.add(data);
		}
	}
	
	public void readSecond(Workbook second){
		
		
		for(Sheet sheet : second){	
			
			//This sheet has no data
			if(sheet.getSheetName().equals("TOTALs")){
				continue;
			}			
			
			boolean first = true;
			
			for(Row row : sheet){				
				
				//Ignore the first row since it has the headers
				if(first){
					first = false;
					continue;
				}
				
				//After the first row...
				secondSet.add(new SecondData(row, sheet.getSheetName()));
			}
		}	
	}
	
	public Workbook openWorkbook(String path){
		try {
			OPCPackage pkg = OPCPackage.open(new File(path));
			Workbook wb = new XSSFWorkbook(pkg);	
			return wb;
		} catch (InvalidFormatException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		
		return null;
	}
}