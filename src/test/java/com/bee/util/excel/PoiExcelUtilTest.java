package com.bee.util.excel;

import java.io.File;
import java.io.FileOutputStream;
import java.io.OutputStream;
import java.util.List;

import com.bee.util.entity.Employee;

/**
 * @Description TODO
 * @author Linfeng
 * @date 2015年12月4日 下午2:04:18
 * 
 **/
public class PoiExcelUtilTest {
	public static void main(String[] args) throws Exception{
		String path = "E:\\emplyee.xlsx";
		List<Employee> list = PoiExcelUtil.readExcelAsObject(path, Employee.class, new String[]{"name","age","sex","salery","addtime"}, true);
		
		File outFile = new File("E:\\生成.xlsx");
		OutputStream outputStream = new FileOutputStream(outFile);
		PoiExcelUtil.writeObjectToExcel(list, outputStream,new String[]{"name","age","sex","addtime"},new String[]{"姓名","年龄","性别","加入时间"});
	}
}
