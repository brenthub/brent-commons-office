package cn.brent.commons.office.excel.handler;

/**
 * 导入导出转换器
 * @param <E> Excel中的字段类型
 * @param <M> 实体中的字段类型
 */
public interface ValueHandler<E,M>{
	
	/**
	 * 导入值转换
	 * @param value
	 * @return
	 */
	M impConvert(E value);
	
	/**
	 * 导出值转换
	 * @param value
	 * @return
	 */
	E expConvert(M value);
	
}