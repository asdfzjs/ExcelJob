package com.zillionfortune.util;

import java.lang.annotation.*;

/**
 * Created by zhangwenjun on 2016/11/18.
 */

@Documented
@Retention(RetentionPolicy.RUNTIME)
@Target(ElementType.TYPE)
public @interface SheetToClass {
    String sheetName();
}
