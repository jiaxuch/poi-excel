package com.jiaxuch.excel.controller;

import com.jiaxuch.excel.model.Teacher;
import com.jiaxuch.excel.service.PoiExcelService;
import com.jiaxuch.poiexcel.SaveDataInterface;
import com.jiaxuch.poiexcel.config.HeaderConfig;
import com.jiaxuch.poiexcel.utils.ExcelParseUtils;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.MediaType;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.ResponseBody;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * @author jiaxuch
 * @data 2020/6/27
 */
@Controller
public class PoiExcelController {

    @Autowired
    PoiExcelService poiExcelService;

    @ResponseBody
    @RequestMapping(value = "/fileUpload",method = RequestMethod.POST, consumes = MediaType.MULTIPART_FORM_DATA_VALUE)
    public List<Teacher> fileUpload(@RequestParam("file") MultipartFile file, HttpServletRequest request) throws Exception {
        Map map = new HashMap();
        map.put("0", "id");
        map.put("1", "age");
        map.put("2", "flag");
        map.put("3", "name");
        map.put("4", "birthday");
        map.put("5", "money");
        SaveDataInterface<Teacher> saveDataInterface = new SaveDataInterface<Teacher>(){
            @Override
            public <T> int save(T t) {
                if(t instanceof Teacher){
                    Teacher teacher = (Teacher) t;
                    return poiExcelService.saveUserInfo(teacher);
                }
                return 0;
            }
        };
        List<Teacher> list = ExcelParseUtils.parseExcelToList(file, map, false,saveDataInterface, Teacher.class);
        return list;
    }

    @RequestMapping(value = "downLoadExcel", method = RequestMethod.GET)
    public void downLoadExcel(HttpServletResponse response) throws Exception {
        List<Teacher> list = poiExcelService.getUserInfo();
        Map<String, HeaderConfig> title = new HashMap<String, HeaderConfig>();
        title.put("0", new HeaderConfig("编号", "id"));
        title.put("1", new HeaderConfig("年龄", "age"));
        title.put("2", new HeaderConfig("标识", "flag"));
        title.put("3", new HeaderConfig("姓名", "name"));
        title.put("4", new HeaderConfig("生日", "birthday"));
        title.put("5", new HeaderConfig("余额", "money"));
        String fileName = "teacher.xlsx";
        ExcelParseUtils.parseListToExcel(list,title,fileName,response);

    }
}
