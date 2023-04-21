package com.localhost.excel.core;

import com.localhost.excel.annotation.ExcelFieldAnnotation;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.joda.time.DateTime;

import java.io.IOException;
import java.io.OutputStream;
import java.lang.reflect.Field;
import java.text.SimpleDateFormat;
import java.util.*;

public class ExcelPortUtils {

    public static <T> void portExcel03(List<T> data, OutputStream outputStream) {
        if (data == null || data.size() == 0) {//健壮性判断
            return;
        }
        //获取数据类型的class对象
        Class<?> classType = data.get(0).getClass();
        //获取类型对象的所有成员变量
        Field[] fields = classType.getDeclaredFields();
        //表头
        List<String> title = new ArrayList<>();
        //根据注解的获取上的字段信息，若没有，则获取成员变量名称，与字段对象封装成map
//        Map<String, Field> map = new LinkedHashMap<>();
        List<Field> fieldList = new ArrayList<>();
        for (int i = 0; i < fields.length; i++) {
            String value = fields[i].getAnnotation(ExcelFieldAnnotation.class).value();
            if (value == null || "".equals(value)) {
                value = fields[i].getAnnotation(ExcelFieldAnnotation.class).fieldName();
            }
            if (value == null || "".equals(value)) {
                value = fields[i].getName();
            }
            title.add(value);
            fieldList.add(fields[i]);
        }
        //创建工作对象
        Workbook workbook = new HSSFWorkbook();
        //03最多能写 65535 行，那么就可以尝试写成两页
        if (data.size() > 50000) {
            Sheet sheet = workbook.createSheet();
            //第一行表头
            Row titleRow = sheet.createRow(0);
            for (int i = 0; i < title.size(); i++) {
                Cell cell = titleRow.createCell(i);
                cell.setCellValue(title.get(i));
            }
            for (int i = 0; i < 50000; i++) {
                Row row = sheet.createRow(i + 1);
                T obj = data.get(i);
                //遍历map，ArrayList是存取有序的，所以能对上表头的信息
                int index = 0;
                for (Iterator<Field> iterator = fieldList.iterator(); iterator.hasNext(); ) {
                    //取出字段类型
                    Field field = iterator.next();
                    Class<?> type = field.getType();
                    try {
                        field.setAccessible(true);
                        Cell cell = row.getCell(index++);
                        Object o = field.get(obj);
                        if (o == null) {
                            continue;
                        }
                        //获取单元格
                        //根据字段类型判断
                        judgeType(type, cell, o);
                    } catch (IllegalAccessException e) {
                        e.printStackTrace();
                    }
                }
            }
            Sheet sheet2 = workbook.createSheet();
            //第一行表头
            Row titleRow2 = sheet.createRow(0);
            for (int j = 0; j < title.size(); j++) {
                Cell cell = titleRow2.createCell(j);
                cell.setCellValue(title.get(j));
            }
            for (int i = 50000; i < data.size(); i++) {
                Row row = sheet2.createRow(i + 1 - 50000);
                T obj = data.get(i);
                //遍历map，ArrayList是存取有序的，所以能对上表头的信息
                int index = 0;
                for (Iterator<Field> iterator = fieldList.iterator(); iterator.hasNext(); ) {
                    //取出字段类型
                    Field field = iterator.next();
                    Class<?> type = field.getType();
                    try {
                        field.setAccessible(true);
                        //获取单元格
                        Cell cell = row.getCell(index++);
                        Object o = field.get(obj);
                        if (o == null) {
                            continue;
                        }
                        //根据字段类型判断
                        judgeType(type, cell, o);
                    } catch (IllegalAccessException e) {
                        e.printStackTrace();
                    }
                }
            }
            try {
                workbook.write(outputStream);
                if (outputStream != null) {
                    outputStream.close();
                }
            } catch (IOException e) {
                e.printStackTrace();
            }
        } else {
            Sheet sheet = workbook.createSheet();
            //第一行表头
            Row titleRow = sheet.createRow(0);
            for (int i = 0; i < title.size(); i++) {
                Cell cell = titleRow.createCell(i);
                cell.setCellValue(title.get(i));
            }
            for (int i = 0; i < data.size(); i++) {
                Row row = sheet.createRow(i + 1);
                T obj = data.get(i);
                //遍历map，ArrayList是存取有序的，所以能对上表头的信息
                int index = 0;
                for (Iterator<Field> iterator = fieldList.iterator(); iterator.hasNext(); ) {
                    //取出字段类型
                    Field field = iterator.next();
                    Class<?> type = field.getType();
                    try {
                        field.setAccessible(true);
                        //获取单元格
                        Cell cell = row.getCell(index++);
                        Object o = field.get(obj);
                        if (o == null) {
                            continue;
                        }
                        //根据字段类型判断
                        judgeType(type, cell, o);
                    } catch (IllegalAccessException e) {
                        e.printStackTrace();
                    }
                }
            }
            try {
                workbook.write(outputStream);
                if (outputStream != null) {
                    outputStream.close();
                }
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }

    public static <T> void portExcel07(List<T> data, OutputStream outputStream) {
        if (data == null || data.size() == 0) {//健壮性判断
            return;
        }
        //获取数据类型的class对象
        Class<?> classType = data.get(0).getClass();
        //获取类型对象的所有成员变量
        Field[] fields = classType.getDeclaredFields();
        //表头
        List<String> title = new ArrayList<>();
        //根据注解的获取上的字段信息，若没有，则获取成员变量名称，与字段对象封装成map
//        Map<String, Field> map = new LinkedHashMap<>();
        List<Field> fieldList = new ArrayList<>();
        for (int i = 0; i < fields.length; i++) {
            String value = fields[i].getAnnotation(ExcelFieldAnnotation.class).value();
            if (value == null || "".equals(value)) {
                value = fields[i].getAnnotation(ExcelFieldAnnotation.class).fieldName();
            }
            if (value == null || "".equals(value)) {
                value = fields[i].getName();
            }
            title.add(value);
            fieldList.add(fields[i]);
        }
        //创建工作对象
        Workbook workbook = new SXSSFWorkbook();
        //03最多能写 65535 行，那么就可以尝试写成两页
        if (data.size() > 50000) {
            Sheet sheet = workbook.createSheet();
            //第一行表头
            Row titleRow = sheet.createRow(0);
            for (int i = 0; i < title.size(); i++) {
                Cell cell = titleRow.createCell(i);
                cell.setCellValue(title.get(i));
            }
            for (int i = 0; i < 50000; i++) {
                Row row = sheet.createRow(i + 1);
                T obj = data.get(i);
                //遍历map，ArrayList是存取有序的，所以能对上表头的信息
                int index = 0;
                for (Iterator<Field> iterator = fieldList.iterator(); iterator.hasNext(); ) {
                    //取出字段类型
                    Field field = iterator.next();
                    Class<?> type = field.getType();
                    try {
                        //获取单元格
                        Cell cell = row.getCell(index++);
                        Object o = field.get(obj);
                        if (o == null) {
                            continue;
                        }
                        judgeType(type, cell, o);
                    } catch (IllegalAccessException e) {
                        e.printStackTrace();
                    }
                }
            }
            Sheet sheet2 = workbook.createSheet();
            //第一行表头
            Row titleRow2 = sheet.createRow(0);
            for (int j = 0; j < title.size(); j++) {
                Cell cell = titleRow2.createCell(j);
                cell.setCellValue(title.get(j));
            }
            for (int i = 50000; i < data.size(); i++) {
                Row row = sheet2.createRow(i + 1 - 50000);
                T obj = data.get(i);
                //遍历map，ArrayList是存取有序的，所以能对上表头的信息
                int index = 0;
                for (Iterator<Field> iterator = fieldList.iterator(); iterator.hasNext(); ) {
                    //取出字段类型
                    Field field = iterator.next();
                    Class<?> type = field.getType();
                    try {
                        //获取单元格
                        Cell cell = row.createCell(index++);
                        Object o = field.get(obj);
                        if (o == null) {
                            continue;
                        }
                        //根据字段类型判断
                        judgeType(type, cell, o);
                    } catch (IllegalAccessException e) {
                        e.printStackTrace();
                    }
                }
            }
            try {
                workbook.write(outputStream);
                // 清除临时文件！
                ((SXSSFWorkbook) workbook).dispose();
                if (outputStream != null) {
                    outputStream.close();
                }
            } catch (IOException e) {
                e.printStackTrace();
            }
        } else {
            Sheet sheet = workbook.createSheet();
            //第一行表头
            Row titleRow = sheet.createRow(0);
            for (int i = 0; i < title.size(); i++) {
                Cell cell = titleRow.createCell(i);
                cell.setCellValue(title.get(i));
            }
            for (int i = 0; i < data.size(); i++) {
                Row row = sheet.createRow(i + 1);
                T obj = data.get(i);
                //遍历map，ArrayList是存取有序的，所以能对上表头的信息
                int index = 0;
                for (Iterator<Field> iterator = fieldList.iterator(); iterator.hasNext(); ) {
                    //取出字段类型
                    Field field = iterator.next();
                    Class<?> type = field.getType();
                    try {
                        field.setAccessible(true);
                        //获取单元格
                        Cell cell = row.createCell(index++);
                        Object o = field.get(obj);
                        if (o == null) {
                            continue;
                        }
                        //根据字段类型判断
                        judgeType(type, cell, o);
                    } catch (IllegalAccessException e) {
                        e.printStackTrace();
                    }
                }
            }

            try {
                workbook.write(outputStream);
                // 清除临时文件！
                ((SXSSFWorkbook) workbook).dispose();
                if (outputStream != null) {
                    outputStream.close();
                }
            } catch (IOException e) {
                e.printStackTrace();
            }
        }

    }
    //判断数据的类型，写入到单元格中
    private static void judgeType(Class<?> type, Cell cell, Object o) {
        //根据字段类型判断
        if (type == Long.class) { //long类型
            Long value = Long.class.cast(o);
            cell.setCellValue(value);
        } else if (type == Integer.class) {//int类型
            Integer value = Integer.class.cast(o);
            cell.setCellValue(value);
        } else if (type == Short.class) {
            Short value = Short.class.cast(o);
            cell.setCellValue(value);
        } else if (type == byte.class) {
            Byte value = Byte.class.cast(o);
            cell.setCellValue(value);
        } else if (type == Date.class) {//日期类型
            Date value = Date.class.cast(o);
            cell.setCellType(HSSFCell.CELL_TYPE_STRING);
            cell.setCellValue(new SimpleDateFormat("yyyy-MM-dd hh:ss:mm").format(value));
        } else if (type == Float.class) { //浮点数
            Float value = Float.class.cast(o);
            cell.setCellValue(value);
        } else if (type == Double.class) {
            Double value = Double.class.cast(o);
            cell.setCellValue(value);
        } else if (type == String.class) {
            String value = String.class.cast(o);
            cell.setCellValue(value);
        } else if (type == Boolean.class) {
            cell.setCellType(HSSFCell.CELL_TYPE_BOOLEAN);
            Boolean value = Boolean.class.cast(o);
            cell.setCellValue(value);
        } else {//为空时候
            //没想好怎么处理，先空着
        }
    }
}
