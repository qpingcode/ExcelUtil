package me.qping.utils.excel.complex;

import lombok.extern.slf4j.Slf4j;
import me.qping.utils.excel.common.Config;
import me.qping.utils.excel.complex.self.*;
import me.qping.utils.excel.handler.WriteHandler;
import org.apache.poi.ss.formula.functions.T;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.RegionUtil;

import java.io.*;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

/**
 * @ClassName ComplexUtil
 * @Description 复杂表头处理类
 * @Author qping
 * @Date 2019/6/25 11:16
 * @Version 1.0
 **/
@Slf4j
public class ComplexUtil {

    // 水平组合，从左往右依次布局
    public static ExcelDiv horizontal(ExcelDiv... excelDivs){
        if(excelDivs == null) return null;

        ExcelDiv first = excelDivs[0];
        for(int i = 1; i < excelDivs.length; i++){
            ExcelDiv second = excelDivs[i];
            first.append(second, ExcelDiv.DIRECTION_RIGHT);
        }
        return first;
    }

    public static ExcelDiv horizontal(List<ExcelDiv> excelDivs){
        if(excelDivs == null) return null;
        ExcelDiv[] array = new ExcelDiv[excelDivs.size()];
        return horizontal(excelDivs.toArray(array));
    }

    // 垂直组合
    public static ExcelDiv vertical(ExcelDiv... excelDivs){
        if(excelDivs == null) return null;

        ExcelDiv first = excelDivs[0];
        for(int i = 1; i < excelDivs.length; i++){
            ExcelDiv second = excelDivs[i];
            first.append(second, ExcelDiv.DIRECTION_BOTTOM);
        }
        return first;
    }

    public static ExcelDiv vertical(List<ExcelDiv> excelDivs){
        if(excelDivs == null) return null;
        ExcelDiv[] array = new ExcelDiv[excelDivs.size()];
        return vertical(excelDivs.toArray(array));
    }

    public static <T> void draw(OutputStream outputStream, ExcelDiv complexHeader, Class<T> clazz, List<T> data, String ext, boolean needSimpleTitle){
        draw(outputStream, complexHeader, clazz, data, ext, needSimpleTitle, false);
    }
    public static <T> void draw(OutputStream outputStream, ExcelDiv complexHeader,
                                Class<T> clazz, List<T> data, String ext, boolean needSimpleTitle ,boolean autoSetColumnWidth){


        Config config = new Config();
        config.initWorkbook(ext);
        config.init(clazz);

        //
        config.setAutoColumn(autoSetColumnWidth);

        // 先画复杂表头
        draw(config, complexHeader);

        // 绘制数据
        WriteHandler writeHandler = new WriteHandler();
        writeHandler.write(config, data, true, needSimpleTitle, complexHeader.getHeight());

        // 输出
        try {
            config.getWorkbook().write(outputStream);
            config.getWorkbook().close();
        } catch (IOException e) {
            e.printStackTrace();
        }

    }

    public static void draw(OutputStream outputStream, ExcelDiv complexHeader){
        Config config = new Config();
        config.initWorkbook("xls");

        draw(config, complexHeader);

        try {
            config.getWorkbook().write(outputStream);
            config.getWorkbook().close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private static void draw(Config config, ExcelDiv complexHeader){

        Workbook workbook = config.getWorkbook();
        Sheet sheet;
        if(workbook.getNumberOfSheets() > 0){
            sheet = workbook.getSheetAt(0);
        }else{
            sheet = workbook.createSheet();
        }

        // 考虑到复杂表头一般不会占用太大的内容，所以每一个 row 都初始化，便于操作
        for(int i = 0; i < complexHeader.getHeight(); i++ ){
            sheet.createRow(i);
        }

        for(me.qping.utils.excel.complex.self.Cell cell : complexHeader.getCellList()){
            org.apache.poi.ss.usermodel.Cell poiCell = sheet.getRow(cell.getPosition().getRow())
                    .createCell(cell.getPosition().getCol());
            poiCell.setCellValue(cell.getValue());

            if(cell.getStyle() != null){
                poiCell.setCellStyle(cell.getStyle().toCellStyle(workbook));
            }
        }




        for(Merge merge: complexHeader.getMergeList()){

            org.apache.poi.ss.usermodel.Cell poiCell = sheet.getRow(merge.getBegin().getRow())
                    .createCell(merge.getBegin().getCol());

            poiCell.setCellValue(merge.getValue());

            CellRangeAddress cellRange = new CellRangeAddress(
                    merge.getBegin().getRow(),
                    merge.getEnd().getRow(),
                    merge.getBegin().getCol(),
                    merge.getEnd().getCol()
            );

            sheet.addMergedRegion(cellRange);

            if(merge.getStyle() != null){
                poiCell.setCellStyle(merge.getStyle().toCellStyle(workbook));
                if(merge.getStyle().isBorder()){
                    RegionUtil.setBorderBottom(BorderStyle.THIN, cellRange, sheet); // 下边框
                    RegionUtil.setBorderLeft(BorderStyle.THIN, cellRange, sheet); // 左边框
                    RegionUtil.setBorderRight(BorderStyle.THIN, cellRange, sheet); // 有边框
                    RegionUtil.setBorderTop(BorderStyle.THIN, cellRange, sheet); // 上边框
                }
            }
        }

        if(config.isAutoColumn()){
            for(int i = 0; i < complexHeader.getWidth(); i++){
                sheet.autoSizeColumn(i);
            }
        }else{
            Iterator<Integer> keyItor = complexHeader.getColWidthMap().keySet().iterator();
            while(keyItor.hasNext()){
                Integer col = keyItor.next();
                int width = 256 * complexHeader.getColWidthMap().get(col) + 184;
                sheet.setColumnWidth(col, width);
            }
        }

        Iterator<Integer> rowItor = complexHeader.getRowHeightMap().keySet().iterator();
        while(rowItor.hasNext()){
            Integer row = rowItor.next();
//            int height = 256 * complexHeader.getRowHeightMap().get(row) + 184;
            sheet.getRow(row).setHeightInPoints(complexHeader.getRowHeightMap().get(row) );
        }

    }

}
