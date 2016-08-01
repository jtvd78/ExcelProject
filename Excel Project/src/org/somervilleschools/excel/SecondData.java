package org.somervilleschools.excel;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

public class SecondData implements Data{
	
	
	
	
	String item;
	String serialNumber;
	String deviceName;
	String currentLocation;
	
	String assetTag;
	String legacyAssetTag;
	String greenAssetTag;
	
	int tag1;
	int tag2;
	int greenTag;
	
	double dateAdded;
	double dateModified;	
	String poNumber;
	String value;	
	String funds;
	String status;
	String note;
	
	String tab;
	int rowNum;

	public SecondData(Row row, String tab) {
		
//		System.out.println(tab + " : " + row.getRowNum());
		
		
//		for(Cell cell : row){
//			System.out.println(Tools.cellTypeToString(cell.getCellType()) + " : " + cell.getColumnIndex() + " : " + cell.toString());
//		}
		
		
		rowNum = row.getRowNum();
		
		this.tab = tab;
		
		this.item = getCellValue(row.getCell(0));
		this.dateAdded = row.getCell(1).getNumericCellValue();
		this.dateModified = row.getCell(2).getNumericCellValue();
		this.legacyAssetTag = row.getCell(3).toString();
		this.assetTag = row.getCell(4).toString();
		this.greenAssetTag = row.getCell(5).toString();
		this.serialNumber = getCellValue(row.getCell(6));		
		this.poNumber = getCellValue(row.getCell(7));	
		this.value = getCellValue(row.getCell(8));
		this.currentLocation = getCellValue(row.getCell(9));
		this.deviceName = row.getCell(10).getStringCellValue();
		this.funds = row.getCell(11).getStringCellValue();
		this.status = row.getCell(12).getStringCellValue();
		this.note = getCellValue(row.getCell(13));	
		
		
		tag1 = Tools.stringToInt(assetTag);
		tag2 = Tools.stringToInt(legacyAssetTag);
		greenTag = Tools.stringToInt(greenAssetTag);
	}
	
	public String getCellValue(Cell cell){
	
		switch(cell.getCellType()){
		
		case Cell.CELL_TYPE_NUMERIC: return String.format("%.0f", cell.getNumericCellValue());
		case Cell.CELL_TYPE_STRING: return cell.toString();
		case Cell.CELL_TYPE_BLANK: return "";
		}
		
		System.out.println("DEFAULT" + cell.getRowIndex() + " : " + tab);
		return "DEFAULT";
		
	}

	@Override
	public String getItem() {
		return item;
	}

	@Override
	public String getSerialNumber() {
		return serialNumber;
	}

	@Override
	public String getDeviceName() {
		return deviceName;
	}

	@Override
	public String getLocation() {
		return currentLocation;
	}

	@Override
	public String getSheetName() {
		return tab;
	}

	@Override
	public int getAssetTagOne() {
		return tag1;
	}

	@Override
	public int getAssetTagTwo() {
		return tag2;
	}

	@Override
	public int getGreenAssetTag() {
		return greenTag;
	}

	@Override
	public String getDataSource() {
		return "Live Asset Inventory";
	}

	@Override
	public int getOriginalRowNumber() {
		return rowNum;
	}	
}