package com.localhost.excel.autoconfigure;

import com.localhost.excel.core.ExcelImportUtils;
import com.localhost.excel.core.ExcelPortUtils;
import org.springframework.context.annotation.Bean;
import org.springframework.context.annotation.Configuration;

@Configuration
public class AnnotationExcelAutoConfig {

    @Bean
    public ExcelImportUtils excelImportUtils(){
        return new ExcelImportUtils();
    }

    @Bean
    public ExcelPortUtils excelPortUtils(){
        return new ExcelPortUtils();
    }

}
