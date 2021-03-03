package me.qping.utils.excel;

import me.qping.utils.excel.common.BeanField;
import me.qping.utils.excel.common.RowConsumer;
import me.qping.utils.excel.common.SheetConsumer;
import me.qping.utils.excel.common.Config;
import me.qping.utils.excel.handler.ReadHandler;
import me.qping.utils.excel.handler.WriteHandler;
import me.qping.utils.excel.handler.big.ReadExcel2007;
import org.apache.poi.ss.usermodel.Sheet;

import java.io.*;
import java.util.*;

/**
 * @ClassName ExcelUtil
 * @Description excel读取封装
 * @Author qping
 * @Date 2019/5/9 17:11
 * @Version 1.0
 **/
public class ExcelUtil {

    private Config config = new Config();

    // 抽离ExcelUtil业务代码，便于阅读
    private WriteHandler writeHandler = new WriteHandler();
    private ReadHandler readHandler = new ReadHandler();

    public Config getConfig(){
        return config;
    }


    public <T> void write(Class<T> clazz, String filePath, Collection<T> data) throws FileNotFoundException {
        String fileExt = "xls";
        if (filePath.endsWith(".xlsx")) {
            fileExt = "xlsx";
        }

        File file = new File(filePath);

        if(!file.exists()){
            file.getParentFile().mkdirs();
        }
        this.write(clazz, new FileOutputStream(file), data, fileExt, true);

    }

    public <T> void write(Class<T> clazz, OutputStream outputStream, Collection<T> data) {
        this.write(clazz, outputStream, data, "xlsx", true);
    }

    public <T> void write(Class<T> clazz, OutputStream outputStream, Collection<T> data, String fileExt) {
        this.write(clazz, outputStream, data, fileExt, true);
    }

    private <T> void write(Class<T> clazz, OutputStream outputStream, Collection<T> data, String fileExt, boolean forceStringType) {
        this.write(clazz, outputStream, data, fileExt, true, null);
    }

    public <T> void write(Class<T> clazz, OutputStream outputStream, Collection<T> data, String fileExt, boolean forceStringType, List<String> ignoreTitles) {
        config.init(clazz);
        config.initWorkbook(fileExt);

        if(ignoreTitles != null && ignoreTitles.size() > 0){
            for(String title: ignoreTitles){
                BeanField ignoreField = null;
                for(BeanField beanField : config.getBeanFields()){
                    if(beanField.getName() != null && beanField.getName().equals(title)){
                        ignoreField = beanField;
                    }
                }

                if(ignoreField != null){
                    config.getBeanFields().remove(ignoreField);
                }
            }
        }

        try {
            writeHandler.write(config, data, forceStringType);
            config.getWorkbook().write(outputStream);
            config.getWorkbook().close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    public List<Map<Integer, Object>> read(String filePath) {
        config.initWorkbook(filePath, filePath.endsWith("xls") ? "xls" : "xlsx");
        Sheet sheet = config.getWorkbook().getSheetAt(config.getSheetNo());
        return readHandler.transferSheetToMapList(
                config.getDataRowBeginNumber(),
                config.isMergeRegionsSeparate(),
                config.getMergeExampleColumnIndex(),
                sheet
        );
    }


    public List<Map<Integer, Object>> read(InputStream inputStream) {
        config.initWorkbook(inputStream);
        Sheet sheet = config.getWorkbook().getSheetAt(config.getSheetNo());
        return readHandler.transferSheetToMapList(
                config.getDataRowBeginNumber(),
                config.isMergeRegionsSeparate(),
                config.getMergeExampleColumnIndex(),
                sheet
        );
    }

    public <T> List<T> read(Class<T> clazz, InputStream inputStream) {

        if(clazz.isAssignableFrom(Map.class)){
            return (List<T>) read(inputStream);
        }

        config.init(clazz);
        config.initWorkbook(inputStream);
        config.initHeader();
        // 读取数据转换为bean
        List<T> list = new ArrayList<>();
        try {
            return readHandler.transferSheetToBeanList(
                    config.getDataRowBeginNumber(),
                    config.getWorkbook().getSheetAt(config.getSheetNo()),
                    config.getClazz(),
                    config.getBeanFields()
            );
        } catch (Exception e) {
            e.printStackTrace();
            return null;
        }
    }

    public <T> void readEachSheet(Class<T> clazz, InputStream inputStream, SheetConsumer<T> sheetConsumer) {
        config.init(clazz);
        config.initWorkbook(inputStream);

        int sheetCount = config.getWorkbook().getNumberOfSheets();
        for (int sheetno = 0; sheetno < sheetCount; sheetno++) {
            config.initHeader(sheetno);
            try {
                List<T> list = readHandler.transferSheetToBeanList(
                        config.getDataRowBeginNumber(),
                        config.getWorkbook().getSheetAt(sheetno),
                        config.getClazz(),
                        config.getBeanFields()
                );

                sheetConsumer.execute(list, config, sheetno, config.getSheetName());
            } catch (Exception e) {
                e.printStackTrace();
            }
        }
    }

    public void readEachSheet(InputStream inputStream, SheetConsumer<Map<Integer,Object>> sheetConsumer) {
        config.initWorkbook(inputStream);

        int sheetCount = config.getWorkbook().getNumberOfSheets();
        for (int sheetno = 0; sheetno < sheetCount; sheetno++) {
            config.initHeader(sheetno);
            try {

                List<Map<Integer, Object>> list = readHandler.transferSheetToMapList(
                        config.getDataRowBeginNumber(),
                        config.isMergeRegionsSeparate(),
                        config.getMergeExampleColumnIndex(),
                        config.getWorkbook().getSheetAt(sheetno)
                );
                sheetConsumer.execute(list, config, sheetno, config.getSheetName());
            } catch (Exception e) {
                e.printStackTrace();
            }
        }
    }

    public static void main(String[] args) {
        ExcelUtil excelUtil = new ExcelUtil();
        excelUtil.sheetNo(0).setDataRowBegin(1).readBig("/Users/qping/Documents/Cloud/work/疾控预警/高淳数据处理/高淳区人民医院电子病历数据.xlsx",
                new RowConsumer<Map<Integer, Object>>(){
                    @Override
                    public void execute(Map<Integer, Object> data, long row, int lastColNum) {


                        System.out.println(data);

                    }
                });
    }

    /**
     * event model
     * @param filePath
     * @param rowConsumer
     */
    public void readBig(String filePath, RowConsumer<Map<Integer, Object>> rowConsumer){

        int sheetNo = config.getSheetNo();
        int dataRowNum = config.getDataRowBeginNumber();

        if(filePath.endsWith(".xlsx")){
            ReadExcel2007 excel2007 = new ReadExcel2007();
            try {
                excel2007.processOneSheet(new FileInputStream(new File(filePath)), sheetNo, dataRowNum, rowConsumer);
            } catch (Exception e) {
                e.printStackTrace();
            }
        }
    }



    // ---------------------------- 参数设置 --------------------------------


    /**
     * 合并单元格的相关配置
     * @param separate  是否拆分合并单元格，当值为true时，每个合并单元格内的cell的值都设为合并单元格的值
     * @param mergeExampleColumnIndex  以某一列为标准，合并单元格跨多行的情况只会解析为一条数据，从0开始
     * @return
     */
    public ExcelUtil mergeStrategy(boolean separate, int mergeExampleColumnIndex){
        this.config.setMergeRegionsSeparate(separate);
        this.config.setMergeExampleColumnIndex(mergeExampleColumnIndex);
        return this;
    }

    /**
     * 设置标题行位置和数据行开始位置
     * @param titleRow
     * @param dataRowBeginNumber
     * @return
     */
    public ExcelUtil setTitleRowAndDataRow(int titleRow, int dataRowBeginNumber){
        this.config.setTitleRow(titleRow);
        this.config.setDataRowBeginNumber(dataRowBeginNumber);
        return this;
    }

    public ExcelUtil setTitleRow(int titleRow){
        this.config.setTitleRow(titleRow);
        return this;
    }

    public ExcelUtil setDataRowBegin(int dataRowBeginNumber){
        this.config.setDataRowBeginNumber(dataRowBeginNumber);
        return this;
    }

    /**
     * 是否首行为标题行
     * @param firstHeader
     * @return
     */
    public ExcelUtil firstHeader(boolean firstHeader) {
        if(firstHeader){
            this.config.setTitleRow(0);
            this.config.setDataRowBeginNumber(1);
        }else{
            this.config.setTitleRow(-1);
            this.config.setDataRowBeginNumber(0);
        }

        return this;
    }

    /**
     * 解析第几个sheet页面，默认为第一个，sheetNo从0开始
     * @param sheetNo
     * @return
     */
    public ExcelUtil sheetNo(int sheetNo) {
        this.config.setSheetNo(sheetNo);
        return this;
    }


    /**
     * 是否拆分合并单元格，当值为true时，每个合并单元格内的cell的值都设为合并单元格的值
     * @param dealMergedCell
     * @return
     */
    public ExcelUtil dealMergedCell(boolean dealMergedCell){
        this.config.setMergeRegionsSeparate(dealMergedCell);
        return this;
    }

}
