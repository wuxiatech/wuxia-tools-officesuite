/*
 * Created on :2017年10月16日
 * Author     :songlin
 * Change History
 * Version       Date         Author           Reason
 * <Ver.No>     <date>        <who modify>       <reason>
 * Copyright 2014-2020 wuxia.gd.cn All right reserved.
 */
package cn.wuxia.common.test;

import cn.wuxia.common.util.reflection.ConvertUtil;
import cn.wuxia.tools.excel.ExportExcelUtil;
import cn.wuxia.tools.excel.bean.ExcelBean;
import jodd.typeconverter.TypeConverterManager;

import java.io.FileOutputStream;
import java.math.BigDecimal;
import java.util.*;

public class ExcelTest {
    public static void main(String[] args) throws Exception {
        test();
    }

    public static void testExport() throws Exception {
        // System.out.println((MAX_ROWS + 1) / MAX_ROWS + 1);
        // Write the output to a file
        long start = System.currentTimeMillis();
        FileOutputStream fileOut1 = new FileOutputStream("/app/workbook1.XLSx");
        // FileOutputStream fileOut2 = new
        // FileOutputStream("c:/workbook2.xls");
        String[] selfields = new String[]{"apply_organ", "business_line", "name", "form_title", "undertake_date", "remark", "accept", "remark_date",
                "user_name", "office_phone"};
        String[] selfieldsName = new String[]{"提出机构", "业务条线", "承办部门名称", "创意名称", "转承办部门日期", "处理意见", "是否认可", "反馈意见日期", "联系人", "电话"};
        List<Map<String, String>> dataList1 = new ArrayList<Map<String, String>>();
        for (int i = 0; i < (16); i++) {
            Map m = new HashMap();
            for (String selfield : selfields) {
                m.put(selfield, i + " 我是仲文中午呢访问访问了福建省辽");
            }
            dataList1.add(m);
        }
        List<Map<String, Object>> dataList2 = new ArrayList<Map<String, Object>>();
        ExcelBean excelBean = new ExcelBean();
        excelBean.setFileName("workbook1.XLSx");
        excelBean.setDataList(dataList1);
        excelBean.setSelfields(selfields);
        excelBean.setSelfieldsName(selfieldsName);
        excelBean.setSheetName("sheet");
        ExportExcelUtil.createExcel(excelBean, fileOut1);

        // excelBean.setDataList(dataList2);
        // ExcelUtil.createExcel(excelBean, fileOut2);

        fileOut1.close();
        // fileOut2.close();
        long end = System.currentTimeMillis();
        System.out.println(("create Excel end, Used " + ((end - start) / 1000) + " s"));
    }

    public static void test() {
        Object convertValue = TypeConverterManager.get().convertType("123", BigDecimal.class);
        System.out.println(convertValue + "" + convertValue.getClass());

        convertValue = ConvertUtil.convert("2017/10/20 22:22:22", Date.class);
        System.out.println(convertValue + "" + convertValue.getClass());
        convertValue = TypeConverterManager.get().convertType("2017/10/20 22:22:22", Date.class);
        System.out.println(convertValue + "" + convertValue.getClass());


        convertValue = TypeConverterManager.get().convertType("2017-10-20", java.sql.Date.class);
        System.out.println(convertValue + "" + convertValue.getClass());
    }
}
