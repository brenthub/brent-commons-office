package junit;

import java.util.List;

import org.apache.commons.lang3.StringUtils;
import org.junit.Test;

import cn.brent.commons.office.excel.BlankRowFilter;
import cn.brent.commons.office.excel.ExcelField;
import cn.brent.commons.office.excel.ExportExcel;
import cn.brent.commons.office.excel.ImportExcel;
import cn.brent.commons.office.excel.handler.NumToStrHandler;

public class ImportExcelTest {

	@Test
	public void testExcel() {
		ImportExcel<MOrderVo> ie = new ImportExcel<MOrderVo>(MOrderVo.class, true, ImportExcelTest.class.getResourceAsStream("/test.xlsx"), 1, 0);

		ie.setBlankRowFilter(new BlankRowFilter<ImportExcelTest.MOrderVo>() {
			@Override
			public boolean isBlankRow(MOrderVo dto) {
				if (StringUtils.isEmpty(dto.getMerId())) {
					return true;
				}
				return false;
			}
		});

		List<MOrderVo> datas = ie.getDatas();

		for (MOrderVo v : datas) {
			System.out.println(v.getMerId() + ":" + v.getMerOrderId());
		}
		System.out.println("end.");

		ExportExcel<MOrderVo> ex = new ExportExcel<MOrderVo>(MOrderVo.class, true, "测试导出");

		ex.setDataList(datas);

		ex.writeFile("target/result.xlsx");
	}

	public static class MOrderVo {

		@ExcelField(sort = 0, handler = NumToStrHandler.class, title = "商家号")
		private String merId;

		@ExcelField(sort = 1, title = "商家订单号")
		private String merOrderId;

		@ExcelField(sort = 2, title = "IP地址")
		private String ip;

		public String getMerId() {
			return merId;
		}

		public void setMerId(String merId) {
			this.merId = merId;
		}

		public String getMerOrderId() {
			return merOrderId;
		}

		public void setMerOrderId(String merOrderId) {
			this.merOrderId = merOrderId;
		}

		public String getIp() {
			return ip;
		}

		public void setIp(String ip) {
			this.ip = ip;
		}

	}
}
