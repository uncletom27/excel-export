package com.meituan.show.settlement.export;

import static java.lang.annotation.RetentionPolicy.RUNTIME;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.Target;

@Retention(RUNTIME)
@Target({ ElementType.FIELD })
public @interface Cell {
    String title() default "";

    CellType type() default CellType.NONE;//暂时无用

    CellStyle style() default CellStyle.NONE;//暂时无用
    
    int order();//序号，用来排序
}
