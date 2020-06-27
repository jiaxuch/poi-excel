package com.jiaxuch.poiexcel;

/**
 * @author jiaxuch
 * @data 2020/6/27
 */
public interface SaveDataInterface<T> {
    /**
     * 保存数据逻辑
     *
     * @param t
     * @param <T>
     * @return
     */
    <T> int  save(T t);
}
