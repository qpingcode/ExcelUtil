package me.qping.utils.excel.handler;

import lombok.extern.slf4j.Slf4j;
import me.qping.utils.excel.common.BeanField;
import me.qping.utils.excel.common.Config;
import me.qping.utils.excel.complex.self.Style;
import me.qping.utils.excel.complex.self.StyleFactory;
import org.apache.poi.ss.usermodel.*;

import java.io.IOException;
import java.io.OutputStream;
import java.text.SimpleDateFormat;
import java.util.Collection;
import java.util.Date;

/**
 * @ClassName WriteHandler
 * @Author qping
 * @Date 2019/5/15 03:05
 * @Version 1.0
 **/
@Slf4j
public class WriteHandler {

    public <T> void write(Config config, Collection<T> data, boolean forceStringType) {
        write(config, data, forceStringType, true, 0);
    }

    /**
     *
     * @param config
     * @param data
     * @param forceStringType
     * @param needTitle
     * @param beginRow  从第几行开始导出，默认为0
     * @param <T>
     */
    public <T> void write(Config config, Collection<T> data, boolean forceStringType, boolean needTitle, int beginRow) {
        Workbook workbook = config.getWorkbook();

        Sheet sheet;
        if(workbook.getNumberOfSheets() > 0){
            sheet = workbook.getSheetAt(0);
        }else{
            sheet = workbook.createSheet();
        }

        int rowIndex = beginRow < 0 ? 0 : beginRow;

        // 输出表头
        if(needTitle){
            Row headerRow = sheet.createRow(rowIndex);

            int headerCol = 0;
            for(BeanField beanField : config.getBeanFields()){

                if(beanField.getName() == null){
                    continue;
                }

                Cell cell = headerRow.createCell(headerCol++);
                cell.setCellValue(beanField.getName());

                if(beanField.getWidth() > -1){
                    sheet.setColumnWidth(headerCol - 1, beanField.getWidth());
                }
            }
            rowIndex++;
        }

        for(T rowData : data){
            Row row = sheet.createRow(rowIndex);

            int colIndex = -1;


            for(BeanField beanField : config.getBeanFields()){
                if(beanField.getName() == null){
                    continue;
                }

                colIndex++;

                try {
                    Object valueObj = beanField.getField().get(rowData);
                    String value = null;

                    if(beanField.getDateformat() != null && valueObj instanceof Date){
                        value = new SimpleDateFormat(beanField.getDateformat()).format(valueObj);
                    }else{
                        value = valueObj == null ? "" : valueObj.toString();
                    }
                    Cell cell = row.createCell(colIndex);
                    cell.setCellValue(value);
                } catch (IllegalAccessException e) {
                    e.printStackTrace();
                }
            }

            rowIndex++;
        }

        // 强制单元格格式为文本类型，防止导出后再导入报错
        if(forceStringType){
            CellStyle textStyle = config.getWorkbook().createCellStyle();
            DataFormat format = config.getWorkbook().createDataFormat();
            textStyle.setDataFormat(format.getFormat("@"));
            int colIndex = -1;
            for(BeanField beanField : config.getBeanFields()){
                colIndex++;
                if(beanField.getField().getType() == String.class){
                    sheet.setDefaultColumnStyle(colIndex, textStyle);
                }
            }


        }


    }

}
