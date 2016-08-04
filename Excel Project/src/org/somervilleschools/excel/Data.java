package org.somervilleschools.excel;

public interface Data {
	
	public String getDataSource();
	public int getOriginalRowNumber();
	
	public String getItem();
	public String getSerialNumber();
	public String getDeviceName();
	public String getLocation();
	public String getSheetName();
	public int getAssetTagOne();
	public int getAssetTagTwo();
	public int getGreenAssetTag();	
	public String getPONumber();
	public int getValue();
	
}