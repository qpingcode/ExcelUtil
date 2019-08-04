package me.qping.utils.excel;

import lombok.Data;
import lombok.extern.slf4j.Slf4j;
import me.qping.utils.excel.common.SheetConsumer;
import me.qping.utils.excel.common.Config;
import me.qping.utils.excel.handler.ReadHandler;
import me.qping.utils.excel.handler.WriteHandler;
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
@Slf4j
@Data
public class ExcelUtil {

    private Config config = new Config();

    // 抽离ExcelUtil业务代码，便于阅读
    private WriteHandler writeHandler = new WriteHandler();
    private ReadHandler readHandler = new ReadHandler();


    public <T> void write(Class<T> clazz, String filePath, Collection<T> data) throws FileNotFoundException {
        String fileExt = "xls";
        if (filePath.endsWith(".xlsx")) {
            fileExt = "xlsx";
        }

        File file = new File(filePath);

        if(!file.exists()){
            file.getParentFile().mkdirs();
        }
        this.write(clazz, new FileOutputStream(file), data, fileExt);

        try {
            config.getWorkbook().close();
        } catch (IOException e) {
        }
    }

    public <T> void write(Class<T> clazz, OutputStream outputStream, Collection<T> data) {
        this.write(clazz, outputStream, data, "xlsx");
    }

    public <T> void write(Class<T> clazz, OutputStream outputStream, Collection<T> data, String fileExt) {
        this.write(clazz, outputStream, data, "xlsx", true);
    }

    public <T> void write(Class<T> clazz, OutputStream outputStream, Collection<T> data, String fileExt, boolean forceStringType) {
        config.init(clazz);
        config.initWorkbook(fileExt);
        writeHandler.write(config, outputStream, data, forceStringType);
    }

    public <T> List<T> read(Class<T> clazz, String filePath) throws FileNotFoundException {
        try (FileInputStream fileInputStream = new FileInputStream(filePath)) {
            return this.read(clazz, fileInputStream);
        } catch (IOException e) {
            e.printStackTrace();
        }
        return null;
    }

    public List<Map<Integer, Object>> read(InputStream inputStream) {
        config.initWorkbook(inputStream);
        return readHandler.transferSheetToMapList(
                config.isFirstHeader(),
                config.isDealMergeRegions(),
                config.getWorkbook().getSheetAt(config.getSheetNo())
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
                    config.isFirstHeader(),
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
                        config.isFirstHeader(),
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
                        config.isFirstHeader(),
                        config.isDealMergeRegions(),
                        config.getWorkbook().getSheetAt(sheetno)

                );
                sheetConsumer.execute(list, config, sheetno, config.getSheetName());
            } catch (Exception e) {
                e.printStackTrace();
            }
        }
    }

    public ExcelUtil firstHeader(boolean firstHeader) {
        this.config.setFirstHeader(firstHeader);
        return this;
    }

    public ExcelUtil sheetNo(int sheetNo) {
        this.config.setSheetNo(sheetNo);
        return this;
    }

    public ExcelUtil dealMergedCell(boolean dealMergedCell){
        this.config.setDealMergeRegions(dealMergedCell);
        return this;
    }

}
