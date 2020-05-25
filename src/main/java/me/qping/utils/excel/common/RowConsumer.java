package me.qping.utils.excel.common;

import java.util.List;

/**
 * @ClassName BigSheetConsumer
 * @Author qping
 * @Date 2019/5/15 03:48
 * @Version 1.0
 **/
public interface RowConsumer<T> {

    public void execute(T data, long row, int lastColNum);
}
