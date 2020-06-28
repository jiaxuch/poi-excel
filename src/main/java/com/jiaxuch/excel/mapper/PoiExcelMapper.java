package com.jiaxuch.excel.mapper;

import com.jiaxuch.excel.model.Teacher;
import org.apache.ibatis.annotations.Mapper;

import java.util.List;

/**
 * @author jiaxuch
 * @data 2020/6/27
 */
@Mapper
public interface PoiExcelMapper {
    /**
     * 查找用户信息
     *
     * @return
     */
    List<Teacher> getUserInfo();

    /**
     * 保存用户数据
     *
     * @param teacher
     * @return
     */
    int saveUserInfo(Teacher teacher);
}
