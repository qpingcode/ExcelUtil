package me.qping.utils;

import static org.junit.Assert.assertTrue;

import me.qping.utils.excel.ExcelUtil;
import org.junit.Test;

import java.io.*;
import java.util.ArrayList;
import java.util.List;

/**
 * Unit test for simple App.
 */
public class AppTest {

    public AppTest() throws FileNotFoundException {
    }

    @Test
    public void exportExcel() throws IOException {


        Drug drug1 = new Drug();
        drug1.setDrugName("test1");
        drug1.setDrugPlatName("test1");
        drug1.setGeneralName("111111");
        drug1.setSpec("1234");


        Drug drug2 = new Drug();
        drug2.setDrugName("test2");
        drug2.setDrugPlatName("test2");
        drug2.setGeneralName("22222");
        drug2.setSpec("3456");


        Drug drug3 = new Drug();
        drug3.setDrugName("test");
        drug3.setDrugPlatName("test1");
        drug3.setGeneralName("3333");
        drug3.setSpec("5678");

        List<Drug> list = new ArrayList<>();
        list.add(drug1);
        list.add(drug2);
        list.add(drug3);

        OutputStream outputStream = new ByteArrayOutputStream();

        new ExcelUtil().write(Drug.class, outputStream, list, "xlsx");

        ((ByteArrayOutputStream) outputStream).writeTo(new FileOutputStream("/Users/qping/test/2.xlsx"));



    }




}
