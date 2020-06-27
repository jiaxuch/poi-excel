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
    List<Teacher> getUserInfo();

    int saveUserInfo(Teacher teacher);
}
