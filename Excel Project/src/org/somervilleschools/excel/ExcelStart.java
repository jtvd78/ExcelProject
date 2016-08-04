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
	
	HashSet<JSSData> jssSet = new HashSet<JSSData>();
	HashSet<OldData> oldSet = new HashSet<OldData>();
	HashSet<NewData> newSet = new HashSet<NewData>();
	
	HashSet<Data> noPONumberValue = new HashSet<Data>();
	
	HashSet<Data> noSerialMatch = new HashSet<Data>();
	HashSet<Data> noSerialNumber = new HashSet<Data>();	
	TreeMap<String, DataMatch> serialMatchMap = new TreeMap<String, DataMatch>();
	
	HashSet<Data> noGreenNumber = new HashSet<Data>();
	HashSet<Data> noGreenMatch = new HashSet<Data>();
	
	TreeMap<Integer, DataMatch> greenMatchMap = new TreeMap<Integer, DataMatch>();
	
	HashSet<Data> noTag = new HashSet<Data>();
	HashSet<Data> noTagMatch = new HashSet<Data>();
	TreeMap<Integer, DataMatch> tagMatchMap = new TreeMap<Integer, DataMatch>();
	
	public static void main(String[] args){	
		new ExcelStart().begin();		
	}
	
	public void begin(){
		
		long start = System.currentTimeMillis();
		
		
	//	Workbook main = openWorkbook("C:\\Users\\jvandort\\Documents\\Excel Project\\Asset_Export.xlsx");
		
		
		Workbook shsData = openWorkbook("C:\\Users\\Justin\\git\\Excel Project\\Excel Project\\Data\\New\\SHS INVENTORY.xlsx");
		Workbook smsData = openWorkbook("C:\\Users\\Justin\\git\\Excel Project\\Excel Project\\Data\\New\\SMS INVENTORY.xlsx");
		Workbook vdvData = openWorkbook("C:\\Users\\Justin\\git\\Excel Project\\Excel Project\\Data\\New\\VDV INVENTORY.xlsx");
		Workbook oldData = openWorkbook("C:\\Users\\Justin\\git\\Excel Project\\Excel Project\\Data\\Old\\Asset_Inventory.xlsx");
		
		//Stops program from throwing errors when it finds an empty cell
		shsData.setMissingCellPolicy(MissingCellPolicy.CREATE_NULL_AS_BLANK);
		smsData.setMissingCellPolicy(MissingCellPolicy.CREATE_NULL_AS_BLANK);
		vdvData.setMissingCellPolicy(MissingCellPolicy.CREATE_NULL_AS_BLANK);
		
		oldData.setMissingCellPolicy(MissingCellPolicy.CREATE_NULL_AS_BLANK);
		
	//	readJSS(main);
		readNewData(shsData);
		readNewData(smsData);
		readNewData(vdvData);
		readOldData(oldData);
		
		System.out.println("Finished reading in " + (System.currentTimeMillis() - start) + " ms");
		
		System.out.println("Start remove no PO Number or Value");
		start = System.currentTimeMillis();
		removeNoPONumberValue();
		System.out.println("Finished removing in " + (System.currentTimeMillis() - start) + " ms");
		
		
		System.out.println("Start Compare");
		start = System.currentTimeMillis();
		compare();
		System.out.println("Finished Comparing in " + (System.currentTimeMillis() - start) + " ms");
		
		start = System.currentTimeMillis();
		
		writeOutput();
		
		System.out.println("Finished writing in " + (System.currentTimeMillis() - start) + " ms");
		
	}
	
	
	
	public void removeNoPONumberValue(){
		
		ArrayList<Data> removeList = new ArrayList<Data>();
		
		for(Data data : oldSet){
			if(data.getPONumber().equals("") && data.getValue() == -1){
				noPONumberValue.add(data);
				removeList.add(data);
			}else{
				System.out.println(data.getPONumber() + " : " + data.getValue());
			}
		}
		
		oldSet.removeAll(removeList);
		
	}
	
	public void writeOutput(){
		Workbook wb = new XSSFWorkbook();
		
		
		

		
		//Write No Serial Number
		Sheet noSerialSheet = createSheet(wb, "No Serial Numbers");
		
		int noSerialCtr = 1;
		for(Data data : noSerialNumber){
			Row row = noSerialSheet.createRow(noSerialCtr);
			
			writeData(data, row);
			
			//Next line
			noSerialCtr++;			
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
		
		//Data Matches by Green Number
		Sheet greenMatchSheet = createSheet(wb, "Green Matches");
		int greenMatchCounter = 1;
		for(DataMatch match : greenMatchMap.values()){
			
			for(Data data : match){
				Row row = greenMatchSheet.createRow(greenMatchCounter);
				
				writeData(data, row);
				
				//Next line
				greenMatchCounter++;
			}
			
			//Next line
			greenMatchCounter++;			
		}
		
		writeSet(noGreenMatch, "No Green Matches", wb);
		writeSet(noGreenNumber, "No Green Numbers", wb);
		writeSet(noPONumberValue, "No PO Number or Value", wb);		
		
		System.out.println("Tag Match Length " + tagMatchMap.size());
		
		//Data Matches by Asset Tag
		Sheet tagMatchSheet = createSheet(wb, "Tag Matches");
		int tagMatchCounter = 1;
		for(DataMatch match : tagMatchMap.values()){
			
			for(Data data : match){
				Row row = tagMatchSheet.createRow(tagMatchCounter);
				
				writeData(data, row);
				
				//Next line
				tagMatchCounter++;
			}
			
			//Next line
			tagMatchCounter++;			
		}
		
		writeSet(noTag, "No Asset Tags", wb);	
		writeSet(noTagMatch, "No Asset Tag match", wb);		
		
		//Write it
		try {			 
			FileOutputStream fileOut = new FileOutputStream("C:\\Users\\Justin\\git\\Excel Project\\Excel Project\\output.xlsx");
			wb.write(fileOut);
			fileOut.close();
			wb.close();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}		
	}
	
	public void writeSet(HashSet<Data> set, String name, Workbook wb){
		Sheet sheet = createSheet(wb, name);
		int ctr = 1;
		for(Data data : set){
			Row row = sheet.createRow(ctr);
			writeData(data, row);
			ctr++;
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
			"Sheet Name",
			"PO Number",
			"Value"
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
		row.createCell(10).setCellValue(data.getPONumber());
		row.createCell(11).setCellValue(data.getValue());
	}
	
	public void compare(){
		
		ArrayList<Data> removeListData = new ArrayList<Data>();
		
		for(Data data : newSet){
			
			String serial = data.getSerialNumber();
			
			//If there's no serial number, add it to the noSerialNumber set
			if(serial.equals("") || serial.equals("N/A")){
				noSerialNumber.add(data);
				continue;
			}
			
			removeListData.add(data);
			
			if(serialMatchMap.containsKey(serial)){
				serialMatchMap.get(serial).addData(data);
			}else{
				serialMatchMap.put(serial, new DataMatch(data));
			}			
		}
		
		newSet.removeAll(removeListData);
		removeListData.clear();
		
		for(Data data : oldSet){
			
			String serial = data.getSerialNumber();
			
			//If there's no serial number, add it to the noSerialNumber set
			if(serial.equals("") || serial.equals("N/A")){
				noSerialNumber.add(data);
				continue;
			}
			
			removeListData.add(data);
			
			if(serialMatchMap.containsKey(serial)){
				serialMatchMap.get(serial).addData(data);
			}else{
				serialMatchMap.put(serial, new DataMatch(data));
			}
		}	
		
		oldSet.removeAll(removeListData);
		removeListData.clear();
		
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
		
		removeList.clear();
		
		//Find matches by Green Tag
		for(Data data : newSet){
			
			int green = data.getGreenAssetTag();
			
			//If there's no serial number, add it to the noSerialNumber set
			if(green == -1){
				noGreenNumber.add(data);
				continue;
			}
			
			if(greenMatchMap.containsKey(green)){
				greenMatchMap.get(green).addData(data);
			}else{
				greenMatchMap.put(green, new DataMatch(data));
			}			
		}
		
		for(Data data : oldSet){
			
			int green = data.getGreenAssetTag();
			
			//If there's no serial number, add it to the noSerialNumber set
			if(green == -1){
				noGreenNumber.add(data);
				continue;
			}
			
			if(greenMatchMap.containsKey(green)){
				greenMatchMap.get(green).addData(data);
			}else{
				greenMatchMap.put(green, new DataMatch(data));
			}			
		}
		
		for(Data data : noSerialNumber){
			
			int green = data.getGreenAssetTag();
			
			//If there's no serial number, add it to the noSerialNumber set
			if(green == -1){
				noGreenNumber.add(data);
				continue;
			}
			
			if(greenMatchMap.containsKey(green)){
				greenMatchMap.get(green).addData(data);
			}else{
				greenMatchMap.put(green, new DataMatch(data));
			}			
		}
		
		for(Data data : noSerialMatch){
			
			int green = data.getGreenAssetTag();
			
			//If there's no serial number, add it to the noSerialNumber set
			if(green == -1){
				noGreenNumber.add(data);
				continue;
			}
			
			if(greenMatchMap.containsKey(green)){
				greenMatchMap.get(green).addData(data);
			}else{
				greenMatchMap.put(green, new DataMatch(data));
			}			
		}
		
		//Find and remove Data with no matches for green Numbers
		//Add move that data to noGreenMatch 
		ArrayList<Integer> removeListInt = new ArrayList<Integer>();
		for(int green : greenMatchMap.keySet()){
			
			DataMatch match = greenMatchMap.get(green);
			
			if(match.size() < 2){
				noGreenMatch.add(match.get(0));
				removeListInt.add(green);
			}
		}
		
		for(Integer remove : removeListInt){
			greenMatchMap.remove(remove);
		}
		
		removeListInt.clear();
		
		//Find matches by serial number
		for(Data data : noGreenMatch){
			
			int tag1 = data.getAssetTagOne();
			int tag2 = data.getAssetTagTwo();
			
			//If there's no serial number, add it to the noSerialNumber set
			if(tag1 == -1 && tag2 == -1){
				noTag.add(data);
				continue;
			}
			
			if(tagMatchMap.containsKey(tag1)){
				tagMatchMap.get(tag1).addData(data);
				
				//Check for 2 just in case
				
				if(tagMatchMap.containsKey(tag2)){
					System.out.println("Hmm");
				}
				
				
			}else{
				
				if(tagMatchMap.containsKey(tag2)){
					tagMatchMap.get(tag2).addData(data);
				}else{
					
					tagMatchMap.put(tag1, new DataMatch(data));
					tagMatchMap.put(tag2, new DataMatch(data));
				}
			}		
		}
		
		//Find matches by serial number
		for(Data data : noGreenNumber){
			
			int tag1 = data.getAssetTagOne();
			int tag2 = data.getAssetTagTwo();
			
			//If there's no serial number, add it to the noSerialNumber set
			if(tag1 == -1 && tag2 == -1){
				noTag.add(data);
				continue;
			}
			
			if(tagMatchMap.containsKey(tag1)){
				tagMatchMap.get(tag1).addData(data);
				
				//Check for 2 just in case
				
				if(tagMatchMap.containsKey(tag2)){
					System.out.println("Hmm");
				}
				
				
			}else{
				
				if(tagMatchMap.containsKey(tag2)){
					tagMatchMap.get(tag2).addData(data);
				}else{
					
					tagMatchMap.put(tag1, new DataMatch(data));
					tagMatchMap.put(tag2, new DataMatch(data));
				}
			}		
		}
		
		//Find and remove Data with no matches for green Numbers
		//Add move that data to noGreenMatch
		for(int green : tagMatchMap.keySet()){
			
			DataMatch match = tagMatchMap.get(green);
			
			if(match.size() < 2){
				noTagMatch.add(match.get(0));
				removeListInt.add(green);
			}
		}
		
		for(Integer remove : removeListInt){
			tagMatchMap.remove(remove);
		}
		
		removeListInt.clear();
		
		
	}
	
	public void readJSS(Workbook main){
		boolean first = true;
		
		for(Row row : main.getSheetAt(0)){
			
			//Ignore the first row since it has the headers
			if(first){
				first = false;
				continue;
			}
			
			JSSData data = new JSSData(row);
			
			//After the first row...
			jssSet.add(data);
		}
	}
	
	public void readOldData(Workbook oldData){
		
		
		for(Sheet sheet : oldData){	
			
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
				oldSet.add(new OldData(row, sheet.getSheetName()));
			}
		}	
	}
	
	public void readNewData(Workbook newData){
		for(Sheet sheet : newData){
			
			boolean first = true;
			
			for(Row row : sheet){				
				
				//Ignore the first row since it has the headers
				if(first){
					first = false;
					continue;
				}
				
				//After the first row...
				newSet.add(new NewData(row, sheet.getSheetName()));
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