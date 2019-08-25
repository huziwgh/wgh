package com.wgh.service;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellRangeAddress;

import net.sf.json.JSONArray;

public class ReadExcel {

	private static final int HEADER_VALUE_TYPE_Z=1;
	private static final int HEADER_VALUE_TYPE_S=2;
	
	public static void main(String[] args) {
		String path="D:\\wgh\\data-transfer\\physicalsystem.xls";
		File file = new File(path);
		ReadExcel read= new ReadExcel();
		JSONArray readExcel = read.readExcel(file,0,2);
		String jsonString= readExcel.toString();
		System.out.println(jsonString);
		read.jsonToFile(jsonString);
	}
	
	
	/***
	 * 判断文件是否为excel
	 * @param file
	 * @return
	 */
	public boolean fileCheck(File file){
		boolean flag = false;
		if(file!=null){
			flag = file.getName().endsWith("xls")||file.getName().endsWith("xlsx");
		}
		return flag;
	}
	
	
	public Row getHeaderRow(Sheet sheet,int headerIndex){
		Row headerRow = null;
		if (sheet !=null){
			headerRow= sheet.getRow(headerIndex);
		}		
		return headerRow;
	}
	//返回所有合并单元格
	public List<CellRangeAddress> getCombineCell(Sheet sheet){
		List<CellRangeAddress> cellList= new ArrayList<>();
		int num = sheet.getNumMergedRegions();
		for (int i = 0; i < num; i++) {
			cellList.add(sheet.getMergedRegion(i));		
		}
		return cellList;
	}
	
	public String getHeaderCellValue(Row headerRow, int cellIndex, int headerType) {
		Cell cell = headerRow.getCell(cellIndex);
		String cellValue = null;
		if (cell != null) {
			if (headerType == HEADER_VALUE_TYPE_Z) {
				cellValue = cell.getRichStringCellValue().getString();
				int l_bracket = cellValue.indexOf("（");
				int r_bracket = cellValue.indexOf("）");
				if (l_bracket==-1){
					l_bracket = cellValue.indexOf("(");
				}
				if (r_bracket==-1){
					r_bracket=cellValue.indexOf(")");
				}
				
				cellValue=cellValue.substring(l_bracket+1,r_bracket);
				
			} else if (headerType == HEADER_VALUE_TYPE_S) {
				cellValue = cell.getRichStringCellValue().getString();
			}
		}
		return cellValue;
	}
	

	public Object getCellValue(Row headerRow, int cellIndex, FormulaEvaluator formula) {
		Cell cell = headerRow.getCell(cellIndex);
		Object cellData = null;
		if (cell != null) {
			switch (cell.getCellType()) {
			case Cell.CELL_TYPE_STRING:
				cellData = cell.getRichStringCellValue().getString();
				break;
			case Cell.CELL_TYPE_BOOLEAN:
				cellData = cell.getBooleanCellValue();
				break;
			case Cell.CELL_TYPE_NUMERIC:
				cellData = cell.getNumericCellValue();
				break;
			case Cell.CELL_TYPE_FORMULA:
				cellData = formula.evaluate(cell).getNumberValue();
				break;
			default:
				cellData = null;
				
			}
		}
		
		return cellData;
	}

	/***
	 * 读取excel文件
	 * 
	 * @param file
	 * @param sheetIndex
	 * @param headerType
	 * @return
	 */
	public JSONArray readExcel(File file,int headerIndex,int headerType){
		List<Map<String,Object>> list = new ArrayList<>();
		if(!fileCheck(file)){
			return null;
		}else {
			try {
				//加载excel表格
				WorkbookFactory wbFactory= new WorkbookFactory();
				Workbook wb= wbFactory.create(file);
				//获取第一个表格
				Sheet sheet= wb.getSheetAt(0);
				//获取表头行
				Row headerRow = getHeaderRow(sheet, headerIndex);
				//获取所有合并单元格
				List<CellRangeAddress> combineCellList = getCombineCell(sheet);
				//读取数据
				FormulaEvaluator formula = wb.getCreationHelper().createFormulaEvaluator();
				for (int i = headerIndex+1; i < sheet.getLastRowNum(); i++) {
					//获取第一行数据
					Row dataRow= sheet.getRow(i);
					//创建存储键值对的map
					Map<String,Object> map= new LinkedHashMap<>();
					//遍历每一行内容放入map中
					for (int j = 0; j < dataRow.getLastCellNum(); j++) {
						//表头为key
						String key= getHeaderCellValue(headerRow,j,headerType);
						//获取表中数据
						Object value= getCellValue(dataRow,j,formula);
						map.put(key, value);						
					}	
					list.add(map);
				}
								
			} catch (Exception e) {
				e.printStackTrace();
			}			
			
		}
		JSONArray jsonArray= JSONArray.fromObject(list);		
		return jsonArray;
	}

	public void jsonToFile(String jsonString){
		
		String path= "D:\\wgh\\data-transfer\\jsonfile\\physicalsystem.txt";
		File file = new File(path);
		//如果文件不存在,创建文件
		if(!file.exists()){
			try {
				file.createNewFile();
			} catch (IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
		}
		//引入输出流
		OutputStream outputStream= null;
		try {
			outputStream= new FileOutputStream(file);//创建输出流
			StringBuilder builder= new StringBuilder();//创建可变字符串变量
			builder.append(jsonString);//拼接字符串
			String context = builder.toString();//将可变字符串转为固定长度字符串
			byte [] bytes = context.getBytes("UTF-8");
			outputStream.write(bytes);
			
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} finally {
			try {
				outputStream.close();
			} catch (IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
		}
				
	}
		
}
