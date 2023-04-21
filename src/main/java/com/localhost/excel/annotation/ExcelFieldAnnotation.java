package com.localhost.excel.annotation;

import org.springframework.context.annotation.Bean;
import org.springframework.core.annotation.AliasFor;

import java.lang.annotation.*;

@Target(ElementType.FIELD)//作用范围：成员变量上
@Retention(RetentionPolicy.RUNTIME) //保留到class字节码阶段
@Documented//可以被抽取到文档中
@Inherited//能被子类继承
public @interface ExcelFieldAnnotation {

    @AliasFor("fieldName")
    String value() default "";

    @AliasFor("value")
    String fieldName() default "";

    int order() default Integer.MAX_VALUE;

}
