package cn.wuxia.tools.excel.exception;

/**
 * 
 * [ticket id]
 * Description of the class 
 * @author songlin.li
 * @ Version : V<Ver.No> <2013年6月30日>
 */
public class ExcelException extends RuntimeException {

    private static final long serialVersionUID = -2623309261327598087L;

    public ExcelException(String msg) {
        super(msg);
    }

    public ExcelException(String message, Throwable cause) {
        super(message, cause);
    }

    public ExcelException(Throwable cause) {
        super(cause);
    }
}
