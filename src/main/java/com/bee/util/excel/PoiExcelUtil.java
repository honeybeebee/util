package com.bee.util.excel;

import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;
import java.io.OutputStream;
import java.lang.reflect.Constructor;
import java.lang.reflect.Field;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * @Description POI操作Excel的工具类
 * @author Linfeng
 * @date 2015年11月20日 下午11:43:47
 * 
 **/
public class PoiExcelUtil {
	
	/**
	 * @Fields EXCEL_SUFFIX_2003 03版本之前的后缀名
	 */
	private final static String EXCEL_SUFFIX_2003 = "xls";
	/**
	 * @Fields EXCEL_SUFFIX_2007 07版本之后的后缀名
	 */
	private final static String EXCEL_SUFFIX_2007 = "xlsx";
	
	/**
	 * @Fields DATE_FORMAT 时间格式化类
	 */
	private final static SimpleDateFormat DATE_FORMAT = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");

	
	/**
	 * 读取excel表格，支持xls和xlsx
	 * 
	 * @param filePath excel文件路径
	 * @param clazz 映射的对象
	 * @param fieldNames excel每一列对应clazz中的属性名
	 * @param hasTitle excel是否含有标题
	 * @return List<T>
	 * @throws Exception
	 */
	@SuppressWarnings("resource")
	public static <T> List<T> readExcelAsObject( String filePath,Class<T> clazz, String[] fieldNames, Boolean hasTitle) throws Exception{
		// Read excel，根据excel的不同版本调用相应的Workbook实例
		Workbook workbook = null;
//		int pointIndex = filePath.lastIndexOf(".");
//		String suffix = filePath.substring(pointIndex+1);
//		InputStream inputStream = new FileInputStream(filePath);
//		if(EXCEL_SUFFIX_2003.equalsIgnoreCase(suffix)){
//			workbook = new HSSFWorkbook(inputStream);
//		}else if(EXCEL_SUFFIX_2007.equalsIgnoreCase(suffix)){
//			workbook = new XSSFWorkbook(inputStream);
//		}
		workbook = WorkbookFactory.create(new File(filePath));
		//获取目标类成员变量的名称和类型
		Map<String, Class<?>> filedMap = getClassFiled(clazz);
		List<T> list = new ArrayList<T>();
		//Read sheet
		Sheet sheet = workbook.getSheetAt(0);
		//Read Row
		//如果有标题从第二行开始
		int firstRowNum = sheet.getFirstRowNum() + (hasTitle ? 1 : 0);
		for (int i = firstRowNum; i <= sheet.getLastRowNum(); i++) {
			Row row = sheet.getRow(i);
			if(row == null){
				continue;
			}
			// 生成实例
			T newInstance = clazz.newInstance();
			for (int j = 0; j < fieldNames.length; j++) {
				//Read Cell
				Cell cell = row.getCell(j);
				if(cell == null){
					continue;
				}
				//获取当前变量名
				String fieldName = fieldNames[j];
				//获取当前变量类型
				Class<?> type = filedMap.get(fieldName);
				//将cell的值转为指定类型
				String cellContent = getCellContent(cell);
				Object cellVal = parseValueWithType(cellContent, type);
				//设置变量值
				Field field = clazz.getDeclaredField(fieldName);
				//取消Java语言的访问检查，以便访问private变量
				field.setAccessible(true);
				field.set(newInstance, cellVal);
			}
			list.add(newInstance);
		}
		return list;
	}
	
	/**获取类各成员变量的名字，类型
	 * @param clazz
	 * @return Map 属性名-属性类型
	 */
	private static <T> Map<String, Class<?>> getClassFiled(Class<T> clazz){
		Map<String,Class<?>> map = new HashMap<String, Class<?>>();
		Field[] fields = clazz.getDeclaredFields();
		for (int i = 0; i < fields.length; i++) {
			Field field = fields[i];
			map.put(field.getName(), field.getType());
		}
		return map;
	}
	
	/**获取单元格的内容
	 * @param cell 单元格
	 * @return String
	 */
	private static String getCellContent(Cell cell) {
		String result = null;
		switch (cell.getCellType()) {
		case Cell.CELL_TYPE_NUMERIC: // 数字
			double val = cell.getNumericCellValue();
			if(DateUtil.isCellDateFormatted(cell)){//日期处理,转成毫秒
				val = DateUtil.getJavaDate(val).getTime();
			}
			//数字没有小数位时只取整数位
			if(val % 1 ==0){
				result = String.valueOf((long)val);
			}else{
				result = String.valueOf(val);
			}
			break;
		case Cell.CELL_TYPE_STRING: // 字符串
			result = cell.getStringCellValue();
			break;
		case Cell.CELL_TYPE_FORMULA: // 公式
			result = cell.getCellFormula();
			break;
		case Cell.CELL_TYPE_BLANK: // 空值
			
			break;
		case Cell.CELL_TYPE_BOOLEAN: // 布尔
			result = String.valueOf(cell.getBooleanCellValue());
			break;
		case Cell.CELL_TYPE_ERROR: // 故障
			
			break;
		default:
			break;
		}
		return result;
	}
	
	/**根据给定的类型，将字符串转化成对应的类型
	 * @param value
	 * @param clazz
	 * @return
	 */
	private static <T> Object parseValueWithType(String value, Class<T> clazz){
		Object result = null;
		//判断是否是基本数据类型
		boolean isPrimitive = clazz.isPrimitive();
		if(isPrimitive){
			//是的话通过对应的包装类来实例化
			try { 
				if (Boolean.TYPE == clazz) {
					result = Boolean.parseBoolean(value);
				} else if (Byte.TYPE == clazz) {
					result = Byte.parseByte(value);
				} else if (Short.TYPE == clazz) {
					result = Short.parseShort(value);
				} else if (Integer.TYPE == clazz) {
					result = Integer.parseInt(value);
				} else if (Long.TYPE == clazz) {
					result = Long.parseLong(value);
				} else if (Float.TYPE == clazz) {
					result = Float.parseFloat(value);
				} else if (Double.TYPE == clazz) {
					result = Double.parseDouble(value);
				} 
			} catch (Exception e) {
				e.printStackTrace();
			}
		}else{
			try {
				//判断是否是Date类型
				if(clazz.isInstance(new Date())){
					result = new Date(Long.parseLong(value));
				}else{
					//调用该类带Sring的构造函数来实例化
					Constructor<T> constructor = clazz.getConstructor(new Class[]{String.class});
					result = constructor.newInstance(value);
				}
			} catch (Exception e) {
				e.printStackTrace();
			}
		}
		return result;
	}
	
	
	/**
	 * 将数据导出为Excel
	 * @param objectList 待导出的类的集合
	 * @param outputStream 输出流
	 * @param fieldNames 类中需要导出的属性名称
	 * @param titles Excel的表格标题，与fieldNames对应
	 * @throws Exception
	 */
	@SuppressWarnings("resource")
	public static <T> void writeObjectToExcel(List<T> objectList,OutputStream outputStream, String[] fieldNames, String[] titles) throws Exception{
		//create workbook，默认使用xlsx格式的excel
		Workbook workbook = new XSSFWorkbook();
		//create sheet
		Sheet sheet = workbook.createSheet();
		//create titles row
		Row titleRow = sheet.createRow(0);
		//create title cell
		for (int i = 0; i < titles.length; i++) {
			Cell cell = titleRow.createCell(i, Cell.CELL_TYPE_STRING);
			cell.setCellValue(titles[i]);
			//设置单元格的宽度
			sheet.setColumnWidth(i, titles[i].length() * 1000);
		}
		//添加表格内容
		for (int i = 0; i < objectList.size(); i++) {
			Object object = objectList.get(i);
			Row objectRow = sheet.createRow(i+1);
			//遍历属性列表
			for (int j = 0; j < fieldNames.length; j++) {
				String fieldName = fieldNames[j];
				String val = getValueByFieldName(object, fieldName);
				Cell cell = objectRow.createCell(j);
				cell.setCellValue(val);
			}
		}
		workbook.write(outputStream);
		workbook.close();
	}
	
	/**根据类变量的名字得到该变量的值
	 * 
	 * @param object
	 * @param fieldName 变量名
	 * @return
	 */
	private static String getValueByFieldName(Object object, String fieldName) {
		StringBuilder result = new StringBuilder();
		try {
			Field field = object.getClass().getDeclaredField(fieldName);
			field.setAccessible(true);
			Object objectVal = field.get(object);
			if(objectVal != null){
				//判断fieldName的类型，时间类需格式化
				Class<?> type = field.getType();
				if(!type.isPrimitive() && type.isInstance(new Date())){
					result.append(DATE_FORMAT.format(objectVal));
				}else{
					result.append(objectVal);
				}
			}
		} catch (Exception e) {
			e.printStackTrace();
		} 
		return result.toString();
	}
	
}
