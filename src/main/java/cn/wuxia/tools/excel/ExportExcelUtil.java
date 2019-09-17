package cn.wuxia.tools.excel;

import cn.wuxia.common.exception.AppServiceException;
import cn.wuxia.common.util.DateUtil;
import cn.wuxia.common.util.ListUtil;
import cn.wuxia.common.util.StringUtil;
import cn.wuxia.common.util.SystemUtil;
import cn.wuxia.common.util.reflection.ReflectionUtil;
import cn.wuxia.tools.excel.annotation.ExcelColumn;
import cn.wuxia.tools.excel.bean.ExcelBean;
import cn.wuxia.tools.excel.exception.ExcelException;
import com.alibaba.excel.support.ExcelTypeEnum;
import org.apache.commons.beanutils.BeanUtils;
import org.apache.commons.io.FilenameUtils;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hssf.util.HSSFColor.HSSFColorPredefined;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.*;
import java.lang.reflect.Field;
import java.net.URLEncoder;
import java.util.Date;
import java.util.List;
import java.util.Map;

/**
 * [ticket id] Description of the class
 *
 * @author songlin.li @ Version : V<Ver.No> <May 17, 2012>
 */
public class ExportExcelUtil {
    private static final Logger logger = LoggerFactory.getLogger(ExportExcelUtil.class);

    private static final int MAX_ROWS = 65535;

    public static Workbook getWorkbookFormExcel(File file) throws FileNotFoundException, IOException {
        Workbook wb = null;

        if (file.getName().toUpperCase().endsWith(".XLS")) {
            wb = createHSSFWorkbook(file);
        } else if (file.getName().toUpperCase().endsWith(".XLSX")) {
            wb = createXSSFWorkbook(file);
        } else {
            throw new AppServiceException(file.getName() + " is not valid excel file.");
        }
        return wb;
    }

    /**
     * @param excelBean
     * @param outputStream
     * @throws Exception
     * @description : map object
     * @author songlin.li
     */
    public static void createExcel(ExcelBean excelBean, OutputStream outputStream) throws Exception {
        long start = System.currentTimeMillis();
        logger.debug("create Excel begin...");
        Workbook wb = null;

        if (excelBean.getFileName().toUpperCase().endsWith(".XLS")) {
            wb = createHSSFWorkbook(excelBean);
        } else if (excelBean.getFileName().toUpperCase().endsWith(".XLSX")) {
            wb = createXSSFWorkbook(excelBean);
        } else {
            throw new Exception("filename is:" + excelBean.getFileName() + " filename should end with .xls or .xlsx");
        }

        long end = System.currentTimeMillis();
        logger.debug("create Excel end, Used " + (end - start) + " ms");
        wb.write(outputStream);
        outputStream.flush();
    }

    private static Workbook createHSSFWorkbook(File file) throws FileNotFoundException, IOException {
        HSSFWorkbook workBook = new HSSFWorkbook(new FileInputStream(file));
        return workBook;
    }

    @SuppressWarnings("unchecked")
    private static Workbook createHSSFWorkbook(ExcelBean excelBean) throws Exception {
        HSSFWorkbook workBook = new HSSFWorkbook();
        List<?> list = excelBean.getDataList();
        int sheetSizes = list.size() / MAX_ROWS + 1;
        if (list.size() % MAX_ROWS == 0) {
            sheetSizes -= 1;
        }
        int listStart = 0;
        for (int s = 1; s < sheetSizes + 1; s++) {
            int listEnd = listStart + MAX_ROWS;
            if (list.size() < listEnd) {
                listEnd = list.size();
            }
            List<?> pageList = list.subList(listStart, listEnd);
            HSSFSheet sheet = workBook.createSheet(excelBean.getSheetName() + s);
            HSSFRow row = sheet.createRow(0);
            String[] head = excelBean.getSelfieldsName();
            for (int i = 0; i < head.length; i++) {
                HSSFCell cell = row.createCell(i);
                cell.setCellValue(head[i]);
                cell.setCellStyle(setHSSFCellStyle(workBook, true));
            }

            CellStyle cs = setHSSFCellStyle(workBook, false);
            for (int i = 0; i < pageList.size(); i++) {
                row = sheet.createRow(i + 1);
                Object object = pageList.get(i);
                if (object instanceof Map) {
                    Map<String, Object> map = (Map<String, Object>) object;
                    for (int j = 0; j < excelBean.getSelfields().length; j++) {
                        String key = excelBean.getSelfields()[j];
                        HSSFCell cell = row.createCell(j);
                        cell.setCellValue(StringUtil.isNotBlank(map.get(key)) ? map.get(key).toString() : "");
                        cell.setCellStyle(cs);
                    }
                } else {
                    for (int j = 0; j < excelBean.getSelfields().length; j++) {
                        String name = excelBean.getSelfields()[j];
                        HSSFCell cell = row.createCell(j);
                        cell.setCellValue(BeanUtils.getSimpleProperty(object, name));
                        cell.setCellStyle(cs);
                    }
                }
            }
            for (int i = 0; i < head.length; i++) {
                sheet.autoSizeColumn(i);
                int colwidth = sheet.getColumnWidth(i);
                if (colwidth < 10 * 2 * 256) {
                    sheet.setColumnWidth(i, 10 * 2 * 256);
                }
            }
            listStart += MAX_ROWS;
        }
        return workBook;
    }

    private static Workbook createXSSFWorkbook(File file) throws FileNotFoundException, IOException {
        return new XSSFWorkbook(new FileInputStream(file));
    }

    @SuppressWarnings("unchecked")
    private static Workbook createXSSFWorkbook(ExcelBean excelBean) throws Exception {
        XSSFWorkbook workBook = new XSSFWorkbook();
        List<?> list = excelBean.getDataList();
        int sheetSizes = list.size() / MAX_ROWS + 1;
        if (list.size() % MAX_ROWS == 0) {
            sheetSizes -= 1;
        }
        int listStart = 0;
        for (int s = 1; s < sheetSizes + 1; s++) {
            int listEnd = listStart + MAX_ROWS;
            if (list.size() < listEnd) {
                listEnd = list.size();
            }
            List<?> pageList = list.subList(listStart, listEnd);
            XSSFSheet sheet = workBook.createSheet(excelBean.getSheetName() + s);
            XSSFRow row = sheet.createRow(0);
            String[] head = excelBean.getSelfieldsName();
            CellStyle cstitle = setXSSFCellStyle(workBook, true);
            for (int i = 0; i < head.length; i++) {
                XSSFCell cell = row.createCell(i);
                cell.setCellStyle(cstitle);
                cell.setCellValue(head[i]);
            }

            CellStyle cs = setXSSFCellStyle(workBook, false);
            for (int i = 0; i < pageList.size(); i++) {
                row = sheet.createRow(i + 1);
                Object object = pageList.get(i);
                if (object instanceof Map) {
                    Map<String, Object> map = (Map<String, Object>) object;
                    for (int j = 0; j < excelBean.getSelfields().length; j++) {
                        String key = excelBean.getSelfields()[j];
                        XSSFCell cell = row.createCell(j);
                        cell.setCellValue(StringUtil.isNotBlank(map.get(key)) ? map.get(key).toString() : "");
                        cell.setCellStyle(cs);
                    }
                } else {
                    for (int j = 0; j < excelBean.getSelfields().length; j++) {
                        String name = excelBean.getSelfields()[j];
                        XSSFCell cell = row.createCell(j);

                        try {
                            cell.setCellValue(BeanUtils.getProperty(object, name));
                        } catch (NoSuchMethodException e) {
                            logger.warn("", e);
                            cell.setCellValue("");
                        }
                        cell.setCellStyle(cs);
                    }
                }
            }
            sheet.setDefaultRowHeight((short) (1.5 * 256)); //设置默认行高，表示2个字符的高度
            sheet.setDefaultColumnWidth(10);//设置默认列宽，实际上回多出2个字符，不知道为什么
            //这只poi组件中的两个方法，需要注意的是，必须先设置列宽然后设置行高，不然列宽没有效果
            for (int i = 0; i < head.length; i++) {
                int bcolwidth = sheet.getColumnWidth(i);
                sheet.autoSizeColumn(i);
                int acolwidth = sheet.getColumnWidth(i);

                System.out.println("列" + i + "，宽度前：" + bcolwidth + "  宽度后：" + acolwidth);
                if (acolwidth < bcolwidth && SystemUtil.IS_OS_LINUX) {
                    System.out.println("列：" + i + "宽:" + bcolwidth);
                    sheet.setColumnWidth(i, bcolwidth);
                }
            }
            listStart = listEnd;
        }
        return workBook;
    }

    private static HSSFFont setHSSFFont(HSSFWorkbook workBook, boolean isTitle) {
        /**
         * Create a font
         */
        HSSFFont font = workBook.createFont();
        /**
         * font color
         */
        // font.setColor(HSSFColor.BLACK.index);
        /**
         * font size
         */
        if (isTitle) {
            font.setFontHeightInPoints((short) 12);
            /**
             * font bold
             */
            font.setBold(true);
        } else {

        }
        /**
         * font
         */
        // font.setFontName("宋体");
        /**
         * FontItalic As Boolean
         */
        font.setItalic(false);
        /**
         * If there is a delete line
         */
        font.setStrikeout(false);
        /**
         * Set superscript, subscript
         */
        font.setTypeOffset(HSSFFont.SS_NONE);
        /**
         * underline
         */
        // font.setUnderline(HSSFFont.U_NONE);
        return font;
    }

    private static HSSFCellStyle setHSSFCellStyle(HSSFWorkbook workBook, boolean isTitle) {
        HSSFCellStyle cellStyle = workBook.createCellStyle();
        cellStyle.setFont(setHSSFFont(workBook, isTitle));
        // cellStyle.setFillForegroundColor(HSSFColor.SKY_BLUE.index);
        // cellStyle.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
        /**
         * set border
         */
        cellStyle.setBorderBottom(BorderStyle.THIN);
        cellStyle.setBorderLeft(BorderStyle.THIN);
        cellStyle.setBorderRight(BorderStyle.THIN);
        cellStyle.setBorderTop(BorderStyle.THIN);
        cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        cellStyle.setAlignment(HorizontalAlignment.CENTER);
        /**
         * border color
         */
        cellStyle.setLeftBorderColor(HSSFColorPredefined.BLACK.getIndex());
        cellStyle.setRightBorderColor(HSSFColorPredefined.BLACK.getIndex());
        cellStyle.setBottomBorderColor(HSSFColorPredefined.BLACK.getIndex());
        cellStyle.setTopBorderColor(HSSFColorPredefined.BLACK.getIndex());
        return cellStyle;
    }

    private static XSSFFont setXSSFFont(XSSFWorkbook workBook, boolean isTitle) {
        /**
         * Create a font
         */
        XSSFFont font = workBook.createFont();
        /**
         * font color
         */
        // font.setColor(HSSFColor.BLACK.index);
        /**
         * font size
         */
        if (isTitle) {
            font.setFontHeightInPoints((short) 12);
            /**
             * font bold
             */
            font.setBold(true);
        } else {

        }
        /**
         * font
         */
        // font.setFontName("宋体");
        /**
         * FontItalic As Boolean
         */
        font.setItalic(false);
        /**
         * If there is a delete line
         */
        font.setStrikeout(false);
        /**
         * Set superscript, subscript
         */
        font.setTypeOffset(XSSFFont.SS_NONE);
        /**
         * underline
         */
        // font.setUnderline(HSSFFont.U_NONE);
        return font;
    }

    private static XSSFCellStyle setXSSFCellStyle(XSSFWorkbook workBook, boolean isTitle) {
        XSSFCellStyle cellStyle = workBook.createCellStyle();
        cellStyle.setFont(setXSSFFont(workBook, isTitle));
        // cellStyle.setFillForegroundColor(HSSFColor.SKY_BLUE.index);
        // cellStyle.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
        /**
         * set border
         */
        cellStyle.setBorderBottom(BorderStyle.THIN);
        cellStyle.setBorderLeft(BorderStyle.THIN);
        cellStyle.setBorderRight(BorderStyle.THIN);
        cellStyle.setBorderTop(BorderStyle.THIN);
        cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        cellStyle.setAlignment(HorizontalAlignment.CENTER);
        /**
         * border color
         */
        cellStyle.setLeftBorderColor((short) 8);
        cellStyle.setRightBorderColor((short) 8);
        cellStyle.setBottomBorderColor((short) 8);
        cellStyle.setTopBorderColor((short) 8);

        return cellStyle;
    }

    /**
     * 按照指定列名导出为excel
     *
     * @param excelName
     * @param sheetName
     * @param excelHeader
     * @param lists
     * @param request
     * @param response
     * @author songlin.li
     */
    @Deprecated
    public static void export(String excelName, String sheetName, String[] excelHeader, List<?> lists, HttpServletRequest request,
                              HttpServletResponse response) {
        try {
            // 创建Excel对象
            HSSFWorkbook wb = new HSSFWorkbook();
            // 创建工作单
            HSSFSheet sheet = wb.createSheet(sheetName);

            // 创建行对象
            HSSFRow row = sheet.createRow(0);

            // 创建标题
            for (int i = 0; i < excelHeader.length; i++) {
                // 创建单元格对象
                HSSFCell cell = row.createCell(i);
                cell.setCellValue(excelHeader[i]);
            }

            // 创建数据
            for (int i = 0; i < lists.size(); i++) {
                row = sheet.createRow(i + 1);
                Object obj = lists.get(i);
                Field[] fields = obj.getClass().getDeclaredFields();

                for (int j = 0; j < fields.length; j++) {
                    Field field = fields[j];
                    if (!field.isAccessible()) {
                        field.setAccessible(true);
                    }
                    Object value = field.get(obj);
                    HSSFCell cell = row.createCell(j);
                    cell.setCellValue((value != null) ? value.toString() : "");
                }

            }

            // 转码
            String agent = request.getHeader("user-agent");
            if (agent.toLowerCase().indexOf("msie") != -1) { // IE
                excelName = URLEncoder.encode(excelName, "UTF-8");
            } else {
                excelName = new String(excelName.getBytes("UTF-8"), "ISO-8859-1");
            }

            response.setHeader("Content-Disposition", "attachment;filename=" + excelName + ".xls");
            // 写出Excel
            wb.write(response.getOutputStream());
            wb.close();
        } catch (Exception e) {
            e.printStackTrace();
        }

    }


    /**
     * 按照指定列名导出为excel
     * 更灵活更多扩展方法参考easyExcel
     * <pre>
     *     ExcelWriter writer = new ExcelWriter(out, ExcelTypeEnum.XLSX, true);
     *
     *             Sheet sheet1 = new Sheet(1, 0);
     *             sheet1.setSheetName("sheet1");
     *             Table table2 = new Table(1);
     * //            table2.setTableStyle(createTableStyle());
     *             table2.setClazz(BaseRowModel.class);
     *             writer.write(exportBeanList, sheet1, table2);
     *
     *             writer.finish();
     * </pre>
     * @param lists
     * @param excelType
     * @param outputStream
     * @author songlin.li
     */
    public static void export(List<?> lists, ExcelTypeEnum excelType, OutputStream outputStream) throws ExcelException {
        try {
            // 创建Excel对象
            Workbook wb = null;

            switch (excelType) {
                case XLS:
                    wb = new HSSFWorkbook();
                    break;
                case XLSX:
                    wb = new XSSFWorkbook();
                    break;
                default:
                    break;
            }
            // 创建工作单
            Sheet sheet = wb.createSheet("sheet1");

            // 创建行对象
            Row row = sheet.createRow(0);

            // 创建标题
            Object o = lists.get(0);
            // 创建单元格对象
            List<ExcelColumn> excelColumn = ReflectionUtil.getAnnotations(o, ExcelColumn.class);
            if (ListUtil.isEmpty(excelColumn)) {
                throw new ExcelException("字段缺少@ExcelColumn注解");
            }
            CellStyle cs = null;

            if (wb instanceof HSSFWorkbook) {
                cs = setHSSFCellStyle((HSSFWorkbook) wb, true);
            } else if (wb instanceof XSSFWorkbook) {
                cs = setXSSFCellStyle((XSSFWorkbook) wb, true);
            }
            for (ExcelColumn excelHeader : excelColumn) {
                Cell cell = row.createCell(excelHeader.colunmIndex());
                cell.setCellValue(excelHeader.columnName());
                cell.setCellStyle(cs);
            }


            List<Field> fields = ReflectionUtil.getAccessibleFields(o.getClass(), false);
            if (wb instanceof HSSFWorkbook) {
                cs = setHSSFCellStyle((HSSFWorkbook) wb, false);
            } else if (wb instanceof XSSFWorkbook) {
                cs = setXSSFCellStyle((XSSFWorkbook) wb, false);
            }
            // 创建数据
            for (int i = 0; i < lists.size(); i++) {
                row = sheet.createRow(i + 1);
                Object obj = lists.get(i);
                for (Field field : fields) {
                    ExcelColumn excelHeader = ReflectionUtil.getAnnotation(field, ExcelColumn.class);
                    if (excelHeader == null) {
                        continue;
                    }
                    Object value = ReflectionUtil.invokeGetterMethod(obj, field.getName());
                    if (null == value) {
                        continue;
                    }
                    if (value instanceof Date) {
                        value = DateUtil.dateToString((Date) value, excelHeader.dateFormat());
                    }
                    Cell cell = row.createCell(excelHeader.colunmIndex());
                    cell.setCellValue((value != null) ? value.toString() : "");
                    cell.setCellStyle(cs);
                }
            }
            // 写出Excel
            wb.write(outputStream);
            wb.close();
        } catch (Exception e) {
            throw new ExcelException("导出处理出错", e);
        }

    }
}
