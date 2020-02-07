package me.qping.utils;

import static org.junit.Assert.assertTrue;

import me.qping.utils.bean.Drug;
import me.qping.utils.excel.ExcelUtil;
import me.qping.utils.excel.complex.ComplexUtil;
import me.qping.utils.excel.complex.self.ExcelDiv;
import me.qping.utils.excel.complex.self.Style;
import me.qping.utils.excel.complex.self.StyleFactory;
import org.junit.Test;

import java.awt.*;
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
    public void complexheaderExportTest() throws IOException {


        Drug drug1 = new Drug();
        drug1.setDrugName("泰诺");
        drug1.setDrugPlatName("复方份麻美敏片");
        drug1.setGeneralName("复方份麻美敏片");
        drug1.setSpec("10g");


        Drug drug2 = new Drug();
        drug2.setDrugName("白加黑");
        drug2.setDrugPlatName("氨酚伪麻美芬片");
        drug2.setGeneralName("氨酚伪麻美芬片");
        drug2.setSpec("12g/片");


        Drug drug3 = new Drug();
        drug3.setDrugName("阿莫西林胶囊");
        drug3.setDrugPlatName("阿莫西林胶囊");
        drug3.setGeneralName("阿莫西林胶囊");
        drug3.setSpec("5g");

        List<Drug> list = new ArrayList<>();
        list.add(drug1);
        list.add(drug2);
        list.add(drug3);

        ExcelDiv excelDiv = ExcelDiv.create(4,3);
        excelDiv
                .merge(0,0,1,3, "这是一个宽4高2的长标题")
                .cell(2,0, "第一个")
                .cell(2,1, "第二个")
                .cell(2,2, "第三个")
                .cell(2,3, "第四个");



        OutputStream outputStream = new ByteArrayOutputStream();

        // 简单导出
//        new ExcelUtil().write(Drug.class, outputStream, list, "xlsx");

        // 复杂表头导出（带数据）
        ComplexUtil.draw(outputStream, excelDiv, Drug.class, list, "xlsx", true);



        ((ByteArrayOutputStream) outputStream).writeTo(new FileOutputStream("/Users/qping/test/2.xlsx"));



    }

    @Test
    public void complexheaderStyleTest() throws IOException {

        Drug drug1 = new Drug();
        drug1.setDrugName("泰诺");
        drug1.setDrugPlatName("复方份麻美敏片");
        drug1.setGeneralName("复方份麻美敏片");
        drug1.setSpec("10g");

        Drug drug2 = new Drug();
        drug2.setDrugName("白加黑");
        drug2.setDrugPlatName("氨酚伪麻美芬片");
        drug2.setGeneralName("氨酚伪麻美芬片");
        drug2.setSpec("12g/片");


        Drug drug3 = new Drug();
        drug3.setDrugName("阿莫西林胶囊");
        drug3.setDrugPlatName("阿莫西林胶囊");
        drug3.setGeneralName("阿莫西林胶囊");
        drug3.setSpec("5g");

        List<Drug> list = new ArrayList<>();
        list.add(drug1);
        list.add(drug2);
        list.add(drug3);

        // 样式
        Style mergeStyle = StyleFactory.FONTBLOD_CENTER_WRAP_BORDER.copy().fontFamily("黑体").fontSize(30).backgroundColor(Color.yellow).width(100);
        Style cellStyle = StyleFactory.FONTBLOD_CENTER.copy().fontFamily("宋体").fontSize(13).fontColor(Color.red);

        ExcelDiv excelDiv = ExcelDiv.create(4,3);
        excelDiv
                .merge(0,0,1,3, "这是一个宽4高2的长标题", mergeStyle)
                .cell(2,0, "第一个", cellStyle)
                .cell(2,1, "第二个", cellStyle)
                .cell(2,2, "第三个", cellStyle)
                .cell(2,3, "第四个", cellStyle);



        OutputStream outputStream = new ByteArrayOutputStream();

        // 简单导出
//        new ExcelUtil().write(Drug.class, outputStream, list, "xlsx");

        // 复杂表头导出（带数据）
        ComplexUtil.draw(outputStream, excelDiv, Drug.class, list, "xlsx", true);



        ((ByteArrayOutputStream) outputStream).writeTo(new FileOutputStream("/Users/qping/test/2.xlsx"));



    }




}
