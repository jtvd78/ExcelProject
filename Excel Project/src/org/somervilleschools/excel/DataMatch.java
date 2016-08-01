package org.somervilleschools.excel;

import java.util.ArrayList;
import java.util.Iterator;

public class DataMatch implements Iterable<Data>{
	
	ArrayList<Data> matchList = new ArrayList<Data>();
	
	public DataMatch(Data data){
		matchList.add(data);
	}
	void addData(Data data){
		matchList.add(data);
	}
	@Override
	public Iterator<Data> iterator() {
		return matchList.iterator();
	}
	
	public int size(){
		return matchList.size();
	}
	
	public Data get(int index){
		return matchList.get(index);
	}	
}