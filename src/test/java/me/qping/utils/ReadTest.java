package me.qping.utils;

import me.qping.utils.bean.SdCompareInspect;
import me.qping.utils.excel.ExcelUtil;
import org.junit.Test;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.util.List;

/**
 * @ClassName ReadTest
 * @Description TODO
 * @Author qping
 * @Date 2021/6/8 10:27
 * @Version 1.0
 **/
public class ReadTest {

    @Test
    public void readTest() throws FileNotFoundException {

        ExcelUtil excelUtil = new ExcelUtil();

        List<SdCompareInspect> list = excelUtil.setTitleRow(0).read(SdCompareInspect.class, new FileInputStream(new File("/Users/qping/test/ch/检验对照模版.xls")));

        System.out.println(list.size());

    }
}
