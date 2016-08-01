package org.somervilleschools.excel;

import org.apache.poi.ss.usermodel.Row;

public class MainData implements Data{	
	
	String model;
	String serialNumber;
	String networkName;
	String location;

	String assetTag;
	String assetNumber;
	
	int tag1;
	int tag2;
	
	String status;
	String assetType;
	String manufacturer;
	String notes1;	
	String notes2;
	
	//Won't need these right now
	String warrantyType;
	String discoverySyncId;
	boolean serviceContract;
	double contractExpirationDate;
	
	int rowNum;
	
	public MainData(Row row){		
		
		assetNumber = row.getCell(0).getStringCellValue();
		assetType = row.getCell(1).getStringCellValue();
		manufacturer = row.getCell(2).getStringCellValue();
		model = row.getCell(3).getStringCellValue();
		status = row.getCell(4).getStringCellValue();
		serialNumber = row.getCell(5).getStringCellValue();
		networkName = row.getCell(6).getStringCellValue();
		location = row.getCell(7).getStringCellValue();
		serviceContract = YNToBoolean(row.getCell(8).getStringCellValue());
		contractExpirationDate = row.getCell(9).getNumericCellValue();
		notes1 = row.getCell(10).getStringCellValue();
		warrantyType = row.getCell(11).getStringCellValue();
		discoverySyncId = row.getCell(12).getStringCellValue();
		assetTag = row.getCell(13).getStringCellValue();
		notes2 = row.getCell(14).getStringCellValue();		
		
		rowNum = row.getRowNum();
		
		tag1 = Tools.stringToInt(notes1);
		tag2 = Tools.stringToInt(assetTag);
	}
	
	public boolean YNToBoolean(String yn){
		if(yn.equals("")){
			return false;
		} else if(yn.equals("Y")){
			return true;
		}else if(yn.equals("N")){
			return false;
		}else 
		
		System.out.println("YN To Boolean problem - '" + yn + "'");
		return false;
	}

	@Override
	public String getItem() {
		return model;
	}

	@Override
	public String getSerialNumber() {
		return serialNumber;
	}

	@Override
	public String getDeviceName() {
		return networkName;
	}
	
	@Override
	public String getLocation(){
		return location;
	}

	@Override
	public String getSheetName() {
		return "Main Sheet";
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
		return -1;
	}
	
	@Override
	public String getDataSource() {
		return "Web Help Desk Asset Export";
	}

	@Override
	public int getOriginalRowNumber() {
		return rowNum;
	}
}