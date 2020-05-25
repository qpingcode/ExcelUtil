package me.qping.utils.excel.handler;

import lombok.extern.slf4j.Slf4j;
import me.qping.utils.excel.common.BeanField;
import me.qping.utils.excel.utils.Util;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;

import java.lang.reflect.Field;
import java.text.SimpleDateFormat;
import java.util.*;

/**
 * @ClassName ReadBigDataHandler
 * @description 大数据页面读取工具
 * @Author qping
 * @Date 2019/5/15 03:04
 * @Version 1.0
 **/
@Slf4j
public class BigReadHandler {

    public <T> List<T> transferSheetToBeanList(int dataRowBegin, Sheet sheet, Class<T> clazz, List<BeanField> beanFields) throws Exception {
        List<T> list = new ArrayList<>();

        Iterator<Row> rowIt = sheet.rowIterator();

        if(dataRowBegin > 0){
            for(int i = 0; i < dataRowBegin; i++){
                if(rowIt.hasNext()){
                    rowIt.next();
                }
            }
        }

        while(rowIt.hasNext()){
            Row row = rowIt.next();
            T obj = clazz.newInstance();

            for(BeanField beanField : beanFields){
                int colIndex = beanField.getIndex();

                if(colIndex == -1){
                    continue;
                }
                Cell cell = row.getCell(colIndex);
                Object cellValue = Util.getCellValue(cell);
                setValue(obj, beanField, cellValue);
            }
            list.add(obj);
        }

        return list;
    }

    private <T>  void setValue(T obj, BeanField beanField, Object cellValue) throws Exception {

        if(cellValue == null){
            return;
        }

        Field field = beanField.getField();
        Class<?> type = field.getType();
        try{
            // 类型为String
            if(type == String.class){

                if(cellValue instanceof Double){
                    beanField.getField().set(obj, String.valueOf(cellValue));
                }else{
                    beanField.getField().set(obj, cellValue);
                }

            }
            // 日期格式化处理
            else if (type == Date.class){
                if(cellValue instanceof String){
                    Date dateVal = new SimpleDateFormat(beanField.getDateformat()).parse((String) cellValue);
                    beanField.getField().set(obj, dateVal);
                }else{
                    beanField.getField().set(obj, cellValue);
                }
            }
            // 类型为Int
            else if(type == Integer.class || type == int.class ){
                if(cellValue instanceof String){
                    int val = Integer.parseInt((String)cellValue);
                    beanField.getField().set(obj, val);
                }
                else if(cellValue instanceof Double){
                    int val = ((Double) cellValue).intValue();
                    beanField.getField().set(obj, val);
                }else{
                    beanField.getField().set(obj, cellValue);
                }
            }else{
                beanField.getField().set(obj, cellValue);
            }

        }catch (Exception ex){
            throw new Exception("值转换错误，属性名称：" + field.getName()
                    + ", 期望类型：" + field.getType().toString()
                    + ", 实际类型：" + cellValue.getClass().getTypeName()
                    + ", 实际值为：" + cellValue
            );
        }


    }


    public List<Map<Integer, Object>> transferSheetToMapList(int dataRowBegin, boolean mergeRegionsSeparate, int mergeExampleColumnIndex, Sheet sheet) {

        List<Map<Integer, Object>> list = new ArrayList<>();
        Iterator<Row> rowIt = sheet.rowIterator();

        if(dataRowBegin > 0){
            for(int i = 0; i < dataRowBegin; i++){
                if(rowIt.hasNext()){
                    rowIt.next();
                }
            }
        }

        int boundRowBegin = -1;
        int boundRowEnd = -1;

        while(rowIt.hasNext()){

            Row row = rowIt.next();
            int rowIndex = row.getRowNum();
            int colCount = row.getLastCellNum();

            if(mergeExampleColumnIndex > -1){
                // 如果需要以某一列为标准，处理横跨多行的合并情况
                CellRangeAddress range = inCellRange(sheet, rowIndex, mergeExampleColumnIndex);

                if(range != null){
                    boundRowBegin = range.getFirstRow();
                    boundRowEnd = range.getLastRow();

                    // 行合并的单元格，仅处理一次
                    if(boundRowBegin != rowIndex){
                        continue;
                    }
                }
            }

            Map<Integer, Object> rowData = new HashMap();
            for(int colIndex = 0; colIndex < colCount; colIndex++){
                if(mergeRegionsSeparate){
                    // 如果需要将每一个合并的单元格拆开，拆开后单元格取值为第一个单元格的内容
                    CellRangeAddress range = mergeRegionsSeparate ? inCellRange(sheet, rowIndex, colIndex) : null;
                    if(range == null){
                        // 不在合并单元格内，或者不开启合并单元格处理
                        Object cellValue = Util.getCellValue(row.getCell(colIndex));
                        rowData.put(colIndex, cellValue);
                    }else{
                        // 如果单元格在合并单元格内，则值为左上角的单元格值
                        int firstCol = range.getFirstColumn();
                        int firstRow = range.getFirstRow();
                        Cell cell = sheet.getRow(firstRow).getCell(firstCol);
                        Object cellValue = Util.getCellValue(cell);
                        rowData.put(colIndex, cellValue);
                    }
                } else {
                   // 如果需要以某一列为标准，处理横跨多行的单元格情况
                    Object cellValue = Util.getCellValue(row.getCell(colIndex));
                    rowData.put(colIndex, cellValue);
                }
            }
            list.add(rowData);
        }

        return list;
    }

    private CellRangeAddress inCellRange(Sheet sheet, int rowIndex, int colIndex) {

        int mergedCount = sheet.getNumMergedRegions();
        for (int i = 0; i < mergedCount; i++) {
            CellRangeAddress range = sheet.getMergedRegion(i);
            if (range.containsRow(rowIndex) && range.containsColumn(colIndex)) {
                return range;
            }
        }
        return null;
    }

}
