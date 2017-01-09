package com.zillionfortune.util;

import java.lang.annotation.*;

/**
 * Created by zhangwenjun on 2016/11/14.
 */

@Documented
@Retention(RetentionPolicy.RUNTIME)
@Target(ElementType.FIELD)
public @interface CellToField {
    int cellIndex();
    ExcelFormatType format() default ExcelFormatType.STRING;
    boolean notNull() default false;
}
