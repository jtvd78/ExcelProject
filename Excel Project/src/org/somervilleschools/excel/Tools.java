package org.somervilleschools.excel;

import org.apache.poi.ss.usermodel.Cell;

public class Tools {

	public static String cellTypeToString(int cellType){
		
		switch(cellType){
		case Cell.CELL_TYPE_BLANK: return "BLANK";
		case Cell.CELL_TYPE_BOOLEAN: return "BOOLEAN";
		case Cell.CELL_TYPE_ERROR: return "ERROR";
		case Cell.CELL_TYPE_FORMULA: return "FORMULA";
		case Cell.CELL_TYPE_NUMERIC: return "NUMERIC";
		case Cell.CELL_TYPE_STRING: return "STRING";
		default: return "DEFAULT_TYPE";
		}
	}
	
	public static int stringToInt(String str){			
		if(str.equals("")){
			return -1;
		}else{
			try{
				return (int)Double.parseDouble(str);
			}catch(NumberFormatException e){
				return -1;
			}
		}
	}	
}
