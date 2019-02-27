/*
* Created on :2017年1月23日
* Author     :songlin
* Change History
* Version       Date         Author           Reason
* <Ver.No>     <date>        <who modify>       <reason>
* Copyright 2014-2020 wuxia.gd.cn All right reserved.
*/
package cn.wuxia.tools.excel.annotation;

import java.lang.annotation.Documented;
import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

import cn.wuxia.common.util.DateUtil.DateFormatter;

@Target({ ElementType.FIELD, ElementType.METHOD })
@Retention(RetentionPolicy.RUNTIME)
@Documented
public @interface ExcelColumn {
  
    /**
     * excel表头列索引，从0开始
     * @author songlin
     * @return
     */
    int colunmIndex();

    /**
     * excel表头列名，如为空则以no为主
     * @author songlin
     * @return
     */
    String columnName() default "";

    /**
     * 如果有值，则此值为默认值将替换excel的列值
     * @author songlin
     * @return
     */
    String defaultValue() default "";
    
    /**
     * 时间指定格式
     * @author songlin
     * @return
     */
    DateFormatter dateFormat() default DateFormatter.FORMAT_YYYY_MM_DD;
    
}
