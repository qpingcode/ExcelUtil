package me.qping.utils.excel.common;

import java.util.Collection;
import java.util.List;

/**
 * @ClassName SheetConsumer
 * @Author qping
 * @Date 2019/5/15 03:48
 * @Version 1.0
 **/
public interface SheetConsumer<T> {

    public void execute(List<T> data, Config config, int sheetNum, String sheetName);
}
