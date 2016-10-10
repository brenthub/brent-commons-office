package cn.brent.office.excel.handler;

import java.text.DecimalFormat;
import java.text.ParseException;

public class NumToStrHandler implements ValueHandler<Object, String>{

	private DecimalFormat format=new DecimalFormat("#"); 
	
	@Override
	public String impConvert(Object value) {
		if(value instanceof Double){
			return format.format(value);
		}else if(value instanceof String){
			return value.toString();
		}else{
			return null;
		}
	}

	@Override
	public Object expConvert(String value) {
		try {
			return format.parse(value);
		} catch (ParseException e) {
			return null;
		}
	}
	
}
