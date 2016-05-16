package com.bee.util.excel;

import java.util.ArrayList;
import java.util.HashSet;
import java.util.List;
import java.util.Set;

/**
 * @Description TODO
 * @author Linfeng
 * @date 2015年12月21日 下午5:27:41
 * 
 **/
public class ExcelTest {

	public static void main(String[] args) throws Exception{
		String filePath = "E:/手机列表.xlsx";
		
		
		List<String> strList = new ArrayList<String>();
		strList.add("tom");
		strList.add("tom");
		strList.add("tom");
		Set<String> set = new HashSet<String>(strList);
		strList = new ArrayList<String>(set);
		System.out.println(strList.size());
	}

}
