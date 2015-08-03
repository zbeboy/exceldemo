package com.b510.excel.util;

import com.b510.common.Common;

public class Util {
	public static String getPostfix(String path){
		if(path == null || Common.EMPTY.equals(path.trim())){
			return Common.EMPTY;
		}
		if(path.contains(Common.POINT)){
			return path.substring(path.lastIndexOf(Common.POINT)+1,path.length());
		}
		return Common.EMPTY;
	}
}
