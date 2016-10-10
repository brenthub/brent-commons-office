package cn.brent.office.excel;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.util.ArrayList;
import java.util.Collections;
import java.util.Comparator;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFFormulaEvaluator;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFFormulaEvaluator;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import cn.brent.office.excel.ExcelField.ExAct;
import cn.brent.office.excel.handler.ValueHandler;

/**
 * 导入Excel文件（支持“XLS”和“XLSX”格式）
 * 
 */
public class ImportExcel<T> {

	protected Logger log = LoggerFactory.getLogger(getClass());

	/**
	 * 工作薄对象
	 */
	protected Workbook wb;

	/**
	 * 工作表对象
	 */
	protected Sheet sheet;

	/**
	 * 标题行号
	 */
	protected final int headerNum;

	/**
	 * 是否为excel2007以上版本
	 */
	protected final boolean isXS;
	
	protected final Class<T> clz;
	
	protected BlankRowFilter<T> blankRowFilter;

	/**
	 * 注解列表（Object[]{ ExcelField, Field/Method, Handler }）
	 */
	private List<Object[]> annoList = new ArrayList<Object[]>();

	/**
	 * 合并单元格value Map
	 */
	protected Map<CellRangeAddress, Object> rangesMap = new HashMap<CellRangeAddress, Object>();
	
	public ImportExcel(Class<T> clz, File file, int headerNum, int sheetIndex) {
		if (file == null) {
			throw new RuntimeException("导入文档为空!");
		} else if (file.getName().toLowerCase().endsWith("xls")) {
			isXS = false;
		} else if (file.getName().toLowerCase().endsWith("xlsx")) {
			isXS = true;
		} else {
			throw new RuntimeException("文档格式不正确!");
		}

		this.headerNum = headerNum;
		this.clz=clz;
		
		try {
			initAnnoList(clz);
			init(new FileInputStream(file), sheetIndex);
		} catch (Exception e) {
			throw new RuntimeException(e);
		}
	}

	public ImportExcel(Class<T> clz, boolean isXs, InputStream is, int headerNum, int sheetIndex) {
		this.isXS = isXs;
		this.headerNum = headerNum;
		this.clz=clz;

		try {
			initAnnoList(clz);
			init(is, sheetIndex);
		} catch (Exception e) {
			throw new RuntimeException(e);
		}
	}

	protected void init(InputStream is, int sheetIndex) {

		if (is == null) {
			throw new RuntimeException("InputStream is null");
		}
		try {
			if (isXS) {
				this.wb = new XSSFWorkbook(is);
			} else {
				this.wb = new HSSFWorkbook(is);
			}
		} catch (IOException e) {
			throw new RuntimeException(e);
		}
		if (this.wb.getNumberOfSheets() < sheetIndex) {
			throw new RuntimeException("文档中没有工作表!");
		}
		this.sheet = this.wb.getSheetAt(sheetIndex);

		int num = sheet.getNumMergedRegions();
		for (int i = 0; i < num; i++) {
			CellRangeAddress cr = sheet.getMergedRegion(i);
			Cell cell = sheet.getRow(cr.getFirstRow()).getCell(cr.getFirstColumn());
			rangesMap.put(cr, getCellValue(cell));
		}

		log.debug("Initialize success.");
	}

	protected Object getCellValue(Cell cell) {
		
		if (cell == null) {
			return null;
		}
		
		try {
			Object val = "";
			if (cell.getCellType() == Cell.CELL_TYPE_NUMERIC) {
				val = cell.getNumericCellValue();
			} else if (cell.getCellType() == Cell.CELL_TYPE_STRING) {
				val = cell.getStringCellValue();
			} else if (cell.getCellType() == Cell.CELL_TYPE_FORMULA) {
				FormulaEvaluator evaluator;
				if (isXS) {
					evaluator = new XSSFFormulaEvaluator((XSSFWorkbook) sheet.getWorkbook());
				} else {
					evaluator = new HSSFFormulaEvaluator((HSSFWorkbook) sheet.getWorkbook());
				}
				int result = evaluator.evaluateFormulaCell(cell);
				if (HSSFCell.CELL_TYPE_ERROR == result) {
					val = cell.getErrorCellValue();
				} else if (result == HSSFCell.CELL_TYPE_NUMERIC) {
					val = cell.getNumericCellValue();
				} else if (result == HSSFCell.CELL_TYPE_STRING) {
					val = cell.getRichStringCellValue().toString();
				}
			} else if (cell.getCellType() == Cell.CELL_TYPE_BOOLEAN) {
				val = cell.getBooleanCellValue();
			} else if (cell.getCellType() == Cell.CELL_TYPE_ERROR) {
				val = cell.getErrorCellValue();
			}
			return val;
		} catch (Exception e) {
			log.error(e.getMessage(),e);
			return null;
		}
	}

	/**
	 * 解析class的注解
	 * 
	 * @param cls
	 * @throws Exception
	 */
	protected void initAnnoList(Class<T> cls) throws Exception {

		Field[] fs = cls.getDeclaredFields();
		for (Field f : fs) {
			ExcelField ef = f.getAnnotation(ExcelField.class);
			if (ef != null && ef.type() != ExAct.exp) {
				Class<?> valType = f.getType();
				ValueHandler<?, ?> handler = null;
				if (ef.handler() != ValueHandler.class) {
					handler = ef.handler().newInstance();
				}
				annoList.add(new Object[] { ef, f, handler, valType });
			}
		}
		// Get annotation method
		Method[] ms = cls.getDeclaredMethods();
		for (Method m : ms) {
			ExcelField ef = m.getAnnotation(ExcelField.class);
			if (ef != null && ef.type() != ExAct.exp) {
				ValueHandler<?, ?> handler = null;
				if (ef.handler() != ValueHandler.class) {
					handler = ef.handler().newInstance();
				}

				Class<?> valType = Class.class;
				if ("get".equals(m.getName().substring(0, 3))) {
					valType = m.getReturnType();
				} else if ("set".equals(m.getName().substring(0, 3))) {
					valType = m.getParameterTypes()[0];
				}
				annoList.add(new Object[] { ef, m, handler, valType });
			}
		}
		// Field sorting
		Collections.sort(annoList, new Comparator<Object[]>() {
			public int compare(Object[] o1, Object[] o2) {
				return new Integer(((ExcelField) o1[0]).sort()).compareTo(new Integer(((ExcelField) o2[0]).sort()));
			};
		});
	}
	
	@SuppressWarnings({ "rawtypes", "unchecked" })
	public List<T> getDatas() {
		
		List<T> dataList = new ArrayList<T>();
		for (int i = this.getDataRowNum(); i < this.getLastDataRowNum(); i++) {
			Row row = this.getRow(i);
			if(row==null){
				continue;
			}
			
			T e;
			try {
				e = clz.newInstance();
			} catch (Exception e1) {
				throw new RuntimeException(e1);
			}
			int column = 0;
			for (Object[] os : annoList) {
				Object val = this.getCellValue(row, column++);
				if (val == null) {
					continue;
				}
				ValueHandler handler=(ValueHandler)os[2];
				Class<?> valType = (Class<?>)os[3];
				try {
					if(os[2]!=null){
						val=handler.impConvert(val);
					}else{
						if (valType == String.class) {
							String s = String.valueOf(val.toString());
							if (StringUtils.endsWith(s, ".0")) {
								val = StringUtils.substringBefore(s, ".0");
							} else {
								val = String.valueOf(val.toString());
							}
						} else if (valType == Integer.class) {
							val = Double.valueOf(val.toString()).intValue();
						} else if (valType == Long.class) {
							val = Double.valueOf(val.toString()).longValue();
						} else if (valType == Double.class) {
							val = Double.valueOf(val.toString());
						} else if (valType == Float.class) {
							val = Float.valueOf(val.toString());
						} else if (valType == Date.class) {
							val = DateUtil.getJavaDate((Double) val);
						}
					}
				} catch (Exception ex) {
					log.info("Get cell value [" + i + "," + column + "] error: " + ex.toString());
					val = null;
				}
				// set entity value
				if (os[1] instanceof Field) {
					Reflections.invokeSetter(e, ((Field) os[1]).getName(), val);
				} else if (os[1] instanceof Method) {
					String mthodName = ((Method) os[1]).getName();
					if ("get".equals(mthodName.substring(0, 3))) {
						mthodName = "set" + StringUtils.substringAfter(mthodName, "get");
					}
					Reflections.invokeMethod(e, mthodName, new Class[] { valType },new Object[] { val });
				}
			}
			
			if(blankRowFilter==null){
				dataList.add(e);
			}else if (!blankRowFilter.isBlankRow(e)){
				dataList.add(e);
			}
		}
		return dataList;
		
	}
	
	/**
	 * 获取单元格值
	 * 
	 * @param row
	 *            获取的行
	 * @param column
	 *            获取单元格列号
	 * @return 单元格值
	 */
	private Object getCellValue(Row row, int column) {
		int rownum = row.getRowNum();
		for (CellRangeAddress cr : rangesMap.keySet()) {
			if (cr.isInRange(rownum, column)) {
				return rangesMap.get(cr);
			}
		}
		return getCellValue(row.getCell(column));
	}
	
	/**
	 * 获取行对象
	 * 
	 * @param rownum
	 * @return
	 */
	private Row getRow(int rownum) {
		return this.sheet.getRow(rownum);
	}

	/**
	 * 获取数据行号
	 * 
	 * @return
	 */
	private int getDataRowNum() {
		return headerNum + 1;
	}

	/**
	 * 获取最后一个数据行号
	 * 
	 * @return
	 */
	private int getLastDataRowNum() {
		return this.sheet.getLastRowNum() + headerNum;
	}

	public void setBlankRowFilter(BlankRowFilter<T> blankRowFilter) {
		this.blankRowFilter = blankRowFilter;
	}
}
