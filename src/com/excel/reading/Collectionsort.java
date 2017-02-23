package com.excel.reading;

import java.util.ArrayList;
import java.util.Collections;
import java.util.Comparator;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Map.Entry;
import java.util.Set;

public class Collectionsort {

	public static void main(String[] args) {
		
		Map<String, String[]> map = new HashMap<String, String[]>();
		
		String[] a={"w","h"};
		String[] b={"f","s"};
		String[] c={"c","a"};
		String[] d={"h","y"};
		String[] e={"t","e"};
		String[] f={"o","k"};
		
		map.put("java", a);
		map.put("C++", b);
		map.put("Java2Novice", c);
		map.put("Unix", d);
		map.put("MAC", e);
		map.put("Why this kolavari", f);
		Set<Entry<String, String[]>> set = map.entrySet();
		List<Entry<String, String[]>> list = new ArrayList<Entry<String, String[]>>(set);
	    Collections.sort( list, new Comparator<Map.Entry<String, String[]>>()
	    {
	    	public int compare( Map.Entry<String, String[]> o1, Map.Entry<String, String[]> o2 )
	    	{return (o1.getValue()[1]).compareTo( o2.getValue()[1] );}
		} );
	    
		for(Map.Entry<String, String[]> entry:list){
		System.out.println(entry.getKey()+" ==== "+entry.getValue()[1]);}

	}

}
