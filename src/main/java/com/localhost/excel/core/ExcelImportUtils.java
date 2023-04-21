package com.localhost.excel.core;

import com.localhost.excel.annotation.ExcelFieldAnnotation;
import javafx.scene.input.DataFormat;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFCellUtil;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.lang.reflect.Field;
import java.util.*;

public class ExcelImportUtils {

    //poi形式，处理excel文件，数据量小
    public static <T> List<T> importExcel(File file, Class<T> obj) {
        String fileName = file.getName();
        String fileType = fileName.substring(fileName.lastIndexOf('.') + 1);
        List<T> data;
        if ("xlsx".equals(fileType)) {
            //07
            data = handleExcel07(file, obj);
        } else {
            //03
            data = handleExcel03(file, obj);
        }
        return data;
    }

    //poi形式，处理excel文件，数据量小
    public static <T> List<T> importExcelOnlyOneSheet(File file, Class<T> obj) {
        String fileName = file.getName();
        String fileType = fileName.substring(fileName.lastIndexOf('.') + 1);
        List<T> data = null;
        if ("xlsx".equals(fileType)) {
            //07
            Workbook workbook = null;
            FileInputStream fileInputStream = null;
            try {
                fileInputStream = new FileInputStream(file);
                workbook = new XSSFWorkbook(fileInputStream);
                //处理多页的数据
                data = handleSheet(workbook.getSheetAt(0), obj);
            } catch (IOException e) {
                //io异常，读取文件的异常
                e.printStackTrace();
            } catch (InstantiationException e) {
                e.printStackTrace();
            } catch (IllegalAccessException e) {
                e.printStackTrace();
            } finally {
                //关闭文件输入流
                if (fileInputStream != null) {
                    try {
                        fileInputStream.close();
                    } catch (IOException e) {
                        e.printStackTrace();
                    }
                }
            }
        } else {
            //03
            Workbook workbook = null;
            FileInputStream fileInputStream = null;
            try {
                fileInputStream = new FileInputStream(file);
                workbook = new HSSFWorkbook(fileInputStream);
                //处理多页的数据
                data = handleSheet(workbook.getSheetAt(0), obj);
            } catch (IOException e) {
                //io异常，读取文件的异常
                e.printStackTrace();
            } catch (InstantiationException e) {
                e.printStackTrace();
            } catch (IllegalAccessException e) {
                e.printStackTrace();
            } finally {
                //关闭文件输入流
                if (fileInputStream != null) {
                    try {
                        fileInputStream.close();
                    } catch (IOException e) {
                        e.printStackTrace();
                    }
                }
            }
        }
        return data;
    }

    //03版本的excel文件
    private static <T> List<T> handleExcel03(File file, Class<T> obj) {
        Workbook workbook = null;
        List<T> objects = new ArrayList<>();
        FileInputStream fileInputStream = null;
        try {
            fileInputStream = new FileInputStream(file);
            workbook = new HSSFWorkbook(fileInputStream);
            int totalOfSheets = workbook.getNumberOfSheets();
            for (int i = 0; i < totalOfSheets; i++) {
                //处理多页的数据
                List<T> data = handleSheet(workbook.getSheetAt(i), obj);
                //添加到最终的集合中
                objects.addAll(data);
            }
        } catch (IOException e) {
            //io异常，读取文件的异常
            e.printStackTrace();
        } catch (InstantiationException e) {
            e.printStackTrace();
        } catch (IllegalAccessException e) {
            e.printStackTrace();
        } finally {
            //关闭文件输入流
            if (fileInputStream != null) {
                try {
                    fileInputStream.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        }
        return objects;
    }

    //07版本的excel文件
    private static <T> List<T> handleExcel07(File file, Class<T> obj) {
        Workbook workbook = null;
        List<T> objects = new ArrayList<>();
        FileInputStream fileInputStream = null;
        try {
            fileInputStream = new FileInputStream(file);
            workbook = new XSSFWorkbook(fileInputStream);
            int totalOfSheets = workbook.getNumberOfSheets();
            for (int i = 0; i < totalOfSheets; i++) {
                //处理多页的数据
                List<T> data = handleSheet(workbook.getSheetAt(i), obj);
                //添加到最终的集合中
                objects.addAll(data);
            }
        } catch (IOException e) {
            //io异常，读取文件的异常
            e.printStackTrace();
        } catch (InstantiationException e) {
            e.printStackTrace();
        } catch (IllegalAccessException e) {
            e.printStackTrace();
        } finally {
            //关闭文件输入流
            if (fileInputStream != null) {
                try {
                    fileInputStream.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        }
        return objects;
    }


    public static <T> List<T> handleSheet(Sheet sheet, Class<T> obj) throws InstantiationException, IllegalAccessException {
        if (sheet == null) {
            //非空判断
            return new ArrayList<>();
        }
        //注解上的字段信息一定要和excel上的对应
        Field[] fields = obj.getDeclaredFields();
        //Map集合，存储excel中每列(列索引)对应字段的成员变量
        Map<Integer, Field> map = new HashMap<>();
        //存反射生成的对象
        List<T> objects = new ArrayList<>();
        //表头列数
        Row titleRow = sheet.getRow(0);
        for (int i = 0; i < titleRow.getPhysicalNumberOfCells(); i++) {
            Cell cell = titleRow.getCell(i);
            String title = cell.getStringCellValue();
            for (int j = 0; j < fields.length; j++) {
                //暴力反射
                fields[i].setAccessible(true);
                //获取字段名称
                String value = fields[i].getAnnotation(ExcelFieldAnnotation.class).value();
                if (value == null || "".equals(value)) {
                    value = fields[i].getAnnotation(ExcelFieldAnnotation.class).fieldName();
                }
                if (value != null) {
                    //标题和字段上的注解对应上
                    if (value.equals(title)) {//先判断注解上的是否给定了字段信息，如果没有，则根据成员对象名称判断
                        map.put(i, fields[i]);
                    } else {
                        value = fields[i].getName();
                        if (value.equals(title)) {
                            map.put(i, fields[i]);
                        }
                    }
                }
            }
        }
        //获取行数
        int rowCount = sheet.getPhysicalNumberOfRows();
        for (int i = 1; i < rowCount; i++) {
            Row row = sheet.getRow(i);
            //根据表头信息索引位置进行数据封装
            //反射生成对象
            T o = obj.newInstance();
            for (Iterator<Map.Entry<Integer, Field>> iterator = map.entrySet().iterator(); iterator.hasNext(); ) {
                Map.Entry<Integer, Field> entry = iterator.next();
                Integer index = entry.getKey();
                Field field = entry.getValue();
                //判断字段类型
                Cell cell = row.getCell(index);
                int cellType = cell.getCellType();
                switch (cellType) {
                    case HSSFCell.CELL_TYPE_STRING: //字符串类型
                        String value = cell.getStringCellValue();
                        field.set(o, value);
                        break;
                    case HSSFCell.CELL_TYPE_NUMERIC: //数字类型
                        if (HSSFDateUtil.isCellDateFormatted(cell)) { //日期
                            Date date = cell.getDateCellValue();
                            field.set(o, date);
                        } else { //单纯的是数字
                            //获取double类型的数据
                            Double numericCellValue = cell.getNumericCellValue();
                            //根据成员变量类型封装对应的数据
                            Class<?> fieldType = field.getType();
                            if (fieldType == Double.class) {
                                field.set(o, numericCellValue);
                            } else if (fieldType == Float.class) {
                                field.set(o, numericCellValue.floatValue());
                            } else if (fieldType == Short.class) {
                                field.set(o, numericCellValue.shortValue());
                            } else if (fieldType == Integer.class) {
                                field.set(o, numericCellValue.intValue());
                            } else if (fieldType == Long.class) {
                                field.set(o, numericCellValue.longValue());
                            } else {//byte
                                field.set(o, numericCellValue.byteValue());
                            }
                        }
                        break;
                    case HSSFCell.CELL_TYPE_BOOLEAN: //布尔值
                        field.set(o, cell.getBooleanCellValue());
                        break;
                    case HSSFCell.CELL_TYPE_BLANK: //空
                        break;
                    default:
                        //记录类型不匹配日志
                        System.out.println("第" + i + "行，第" + index + "列的" + cell + "类型无法匹配");
                        break;
                }

            }
            objects.add(o);
        }
        return objects;
    }
}
