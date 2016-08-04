package org.somervilleschools.excel;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

public class NewData implements Data{
	
	int row;
	String location;
	String item;
	String name;
	String ip;
	String serial;
	String assetTagGreen;
	String assetTagYellow;
	String assetTagWhite;
	String funding;
	String poNumber;
	int value;

	int tag1;
	int tag2;
	int greentag;
	
	String sheet;
	
	public NewData(Row row, String sheetName) {
		
		System.out.println(row.getRowNum() + " : " + sheetName);
		
		this.row = row.getRowNum();
		this.sheet = sheetName;
		
		this.item = row.getCell(0).getStringCellValue();
		this.name = getCellValue(row.getCell(1));
		this.ip = getCellValue(row.getCell(2));
		this.serial = getCellValue(row.getCell(3));
		this.greentag = Tools.stringToInt(getCellValue(row.getCell(4)));
		this.tag1 = Tools.stringToInt(getCellValue(row.getCell(5)));
		this.tag2 = Tools.stringToInt(getCellValue(row.getCell(6)));
		this.location = getCellValue(row.getCell(7));
		this.funding = getCellValue(row.getCell(8));
		this.poNumber = getCellValue(row.getCell(9));
		this.value = Tools.stringToInt(getCellValue(row.getCell(10)));
		
		
		
	}
	
	public String getCellValue(Cell cell){
		
		switch(cell.getCellType()){
		
		case Cell.CELL_TYPE_NUMERIC: return String.format("%.0f", cell.getNumericCellValue());
		case Cell.CELL_TYPE_STRING: return cell.toString();
		case Cell.CELL_TYPE_BLANK: return "";
		}
		
		System.out.println("DEFAULT" + cell.getRowIndex() + " : " + location);
		return "DEFAULT";
		
	}

	@Override
	public String getDataSource() {
		return "New Data";
	}

	@Override
	public int getOriginalRowNumber() {
		return row;
	}

	@Override
	public String getItem() {
		return item;
	}

	@Override
	public String getSerialNumber() {
		return serial;
	}

	@Override
	public String getDeviceName() {
		return name;
	}

	@Override
	public String getLocation() {
		return location;
	}

	@Override
	public String getSheetName() {
		return sheet;
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
		return greentag;
	}

	@Override
	public String getPONumber() {
		return poNumber;
	}

	@Override
	public int getValue() {
		return value;
	}

}
