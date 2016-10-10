package cn.brent.office.excel;

public interface BlankRowFilter<T> {

	boolean isBlankRow(T dto);
	
}
