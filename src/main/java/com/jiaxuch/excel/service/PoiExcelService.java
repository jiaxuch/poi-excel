package com.jiaxuch.excel.service;

import com.jiaxuch.excel.mapper.PoiExcelMapper;
import com.jiaxuch.excel.model.Teacher;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;
import org.springframework.transaction.annotation.Transactional;

import java.util.List;

/**
 * @author jiaxuch
 * @data 2020/6/27
 */
@Service
@Transactional(rollbackFor = Exception.class)
public class PoiExcelService {

    @Autowired
    PoiExcelMapper poiExcelMapper;

    public List<Teacher> getUserInfo() {
        return poiExcelMapper.getUserInfo();
    }

    public int saveUserInfo(Teacher teacher) {
        return poiExcelMapper.saveUserInfo(teacher);
    }
}
