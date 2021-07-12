

/**
 * @author: zhliang
 * @version: 1.0
 * @since 2019-07-29
 */
package me.qping.utils.bean;

import java.io.Serializable;

import lombok.Data;
import me.qping.utils.excel.anno.Excel;


@Data
public class SdCompareInspect implements Serializable {

    private Long resultId;

    @Excel(name = "项目编码")
    private String code;
    @Excel(name = "项目名称")
    private String name;
    @Excel(name = "检测方法")
    private String way;
    @Excel(name = "标尺")
    private String staff;
    @Excel(name = "标本代码")
    private String specimenCode;
    @Excel(name = "标本名称")
    private String specimenName;
    /**检验类型编码*/
    @Excel(name = "检验类型编码")
    private String typeCode;
    /**备注说明*/
    @Excel(name = "备注说明")
    private String remark;
    /**被比对的目标编码*/
    @Excel(name = "目标编码")
    private String targetCode;

}