package cn.wuxia.tools.excel.bean;

import java.util.List;

/**
 * [ticket id] Description of the class
 * 
 * @author songlin.li @ Version : V<Ver.No> <May 17, 2012>
 */
public class ExcelBean {
    private String fileName;

    private String sheetName;

    private String[] selfields;

    private String[] selfieldsName;

    private List<?> dataList;

    public ExcelBean() {
    }

    public ExcelBean(String fileName) {
        this.fileName = fileName;
    }

    /**
     * @return
     */
    public String getFileName() {
        return fileName;
    }

    /**
     * @param String
     */
    public void setFileName(String fileName) {
        this.fileName = fileName;
    }

    public String getSheetName() {
        return sheetName;
    }

    public void setSheetName(String sheetName) {
        this.sheetName = sheetName;
    }

    public String[] getSelfields() {
        return selfields;
    }

    public void setSelfields(String[] selfields) {
        this.selfields = selfields;
    }

    public String[] getSelfieldsName() {
        return selfieldsName;
    }

    public void setSelfieldsName(String[] selfieldsName) {
        this.selfieldsName = selfieldsName;
    }

    public List<?> getDataList() {
        return dataList;
    }

    public void setDataList(List<?> dataList) {
        this.dataList = dataList;
    }
}
