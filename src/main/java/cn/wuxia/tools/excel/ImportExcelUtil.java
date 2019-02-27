package cn.wuxia.tools.excel;

import cn.wuxia.common.exception.ValidateException;
import cn.wuxia.common.util.DateUtil;
import cn.wuxia.common.util.ListUtil;
import cn.wuxia.common.util.StringUtil;
import cn.wuxia.common.util.ValidatorUtil;
import cn.wuxia.common.util.reflection.ReflectionUtil;
import cn.wuxia.tools.excel.annotation.ExcelColumn;
import cn.wuxia.tools.excel.exception.ExcelException;
import com.google.common.collect.Lists;
import jodd.typeconverter.TypeConverterManager;
import org.apache.commons.io.FileUtils;
import org.apache.commons.io.IOUtils;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.text.DecimalFormat;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

/**
 * [ticket id] Description of the class
 *
 * @author songlin.li @ Version : V<Ver.No> <May 17, 2012>
 */
public class ImportExcelUtil {
    private static final Logger logger = LoggerFactory.getLogger(ImportExcelUtil.class);

    private static final int MAX_ROWS = 65535;


    /**
     * 导入xls文档
     *
     * @param file
     * @param clazz 需要导入的xls文件
     * @author songlin.li
     */
    public static <T> List<T> importExcel(File file, Class<T> clazz) throws ExcelException {
        if (file != null && file.exists() && !file.isDirectory()) {
            try {
                return importExcel(FileUtils.openInputStream(file), clazz);
            } catch (IOException | EncryptedDocumentException | InvalidFormatException e) {
                logger.error("", e);
                throw new ExcelException("文件不存在！", e);
            }
        } else {
            throw new ExcelException("文件不存在！");
        }
    }

    /**
     * 导入xls文档
     *
     * @param inputStream
     * @param clazz       需要导入的xls文件
     * @throws IOException
     * @throws InvalidFormatException
     * @throws EncryptedDocumentException
     * @author songlin.li
     */
    public static <T> List<T> importExcel(InputStream inputStream, Class<T> clazz)
            throws ExcelException, EncryptedDocumentException, InvalidFormatException, IOException {
        // 创建最终返回的集合
        List<T> lists = new ArrayList<>();
        // 获得工作薄
        Workbook wb = WorkbookFactory.create(inputStream);
        // 获得第一个工作单
        Sheet sheet = wb.getSheetAt(0);
        // 获得行迭带器
        Iterator<Row> rows = sheet.iterator();

        int index = 0;
        List<String> exceptions = Lists.newArrayList();
        while (rows.hasNext()) {
            Row row = rows.next();
            if (index > 0 && !isRowEmpty(row)) {
                // 导入文件不需要标题
                try {
                    T obj = getRowObject(row, clazz);
                    ValidatorUtil.validate(obj);
                    lists.add(obj);
                } catch (ValidateException e) {
                    logger.error("", e);
                    exceptions.add("第" + (index + 1) + "行，" + e.getMessage());
                } catch (Exception e) {
                    logger.error("", e);
                    exceptions.add("第" + (index + 1) + "行，" + e.getMessage());
                }
            }
            index++;
        }
        if (ListUtil.isNotEmpty(exceptions)) {
            throw new ExcelException(StringUtil.join(exceptions, "\t\n"));
        }
        IOUtils.closeQuietly(inputStream);
        return lists;
    }

    /**
     * 校验是否为空行
     *
     * @param row
     * @return
     */
    public static boolean isRowEmpty(Row row) {
        for (int c = row.getFirstCellNum(); c < row.getLastCellNum(); c++) {
            Cell cell = row.getCell(c);
            if (cell != null && CellType.BLANK.compareTo(cell.getCellTypeEnum()) != 0)
                return false;
        }
        return true;
    }


    public static <T> T getRowObject(Row row, Class<T> clazz) throws ExcelException, InstantiationException, IllegalAccessException {
        // 创建集合用于保存一行的单元格数据
        List<String> exceptions = Lists.newArrayList();
        // 创建对象,注入数据
        T obj = clazz.newInstance();
        List<Field> fields = ReflectionUtil.getAccessibleFields(clazz);
        for (Field field : fields) {
            if (field.isAccessible() && !StringUtil.equals("serialVersionUID", field.getName())) {
                ExcelColumn excelHead = ReflectionUtil.getAnnotation(field, ExcelColumn.class);
                if (excelHead == null) {
                    Method method = ReflectionUtil.getGetterMethodByPropertyName(obj, field.getName());
                    if (null != method) {
                        excelHead = ReflectionUtil.getAnnotation(method, ExcelColumn.class);
                    }
                }
                if (excelHead != null) {
                    try {
                        Cell cell = row.getCell(excelHead.colunmIndex());
                        if (cell != null) {
                            /**
                             * 根据field类型转换相应值
                             */
                            setFieldCellValue(cell, obj, field);
                        }
                    } catch (ExcelException e) {
                        exceptions.add("第" + (excelHead.colunmIndex() + 1) + "列，表头为：" + excelHead.columnName() + "，赋值属性名：" + field.getName() + "，值："
                                + ReflectionUtil.getFieldValue(obj, field.getName()) + "【详细错误】" + e.getMessage());
                    }
                } else {
                    logger.info("跳过：" + field.getName() + "赋值");
                }
            } else {
                logger.info("跳过不可访问的属性：" + field.getName());
            }

        }
        if (ListUtil.isNotEmpty(exceptions)) {
            throw new ExcelException(StringUtil.join(exceptions, "；"));
        }
        return obj;
    }


    /**
     * 得到一个单元格内的值, 并根据特定类型赋值
     *
     * @param cell
     * @param bean
     * @param field
     * @return
     * @author songlin.li
     */
    public static void setFieldCellValue(Cell cell, Object bean, Field field) {
        if (cell == null) {
            return;
        }
        Object fieldValue = null;
        CellType cellType = cell.getCellTypeEnum();
        String fieldType = field.getType().getName();
        switch (cellType) {
            case STRING:
                fieldValue = StringUtil.trim(cell.getStringCellValue());
                break;

            case NUMERIC:
                Double value = cell.getNumericCellValue();
                // 读取日期进行
                if (org.apache.poi.ss.usermodel.DateUtil.isCellDateFormatted(cell)) {
                    ExcelColumn an = field.getAnnotation(ExcelColumn.class);
                    java.util.Date value2 = org.apache.poi.ss.usermodel.DateUtil.getJavaDate((Double) value);
                    fieldValue = DateUtil.dateToString(value2, an.dateFormat());
                } else if (fieldType.equals("java.lang.Integer") || fieldType.equals("int") || fieldType.equals("java.lang.Long")
                        || fieldType.equals("long")) {
                    DecimalFormat df = new DecimalFormat("#");// 转换成整型
                    fieldValue = df.format(value);
                } else if (fieldType.equals("java.lang.String")) {
                    cell.setCellType(CellType.STRING);
                    fieldValue = StringUtil.trim(cell.getStringCellValue());
                } else
                    fieldValue = value;
                break;
            case BOOLEAN:
                fieldValue = cell.getBooleanCellValue();
                break;

            case FORMULA:
                fieldValue = cell.getArrayFormulaRange().formatAsString();
                break;
            case BLANK:
                fieldValue = "";
                break;
        }
        field.setAccessible(true);
        if (fieldValue != null) {
            try {
                Object convertValue = TypeConverterManager.get().convertType(fieldValue, field.getType());
                field.set(bean, convertValue);
            } catch (IllegalArgumentException | IllegalAccessException e) {
                logger.warn(e.getMessage());
                throw new ExcelException("无法赋值，原因是字段类型：" + field.getType() + "，值类型：" + fieldValue.getClass() + "，值：" + fieldValue);
            }
        }
    }
}
