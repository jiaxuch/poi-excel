package com.jiaxuch.poiexcel.utils;

import com.jiaxuch.poiexcel.SaveDataInterface;
import com.jiaxuch.poiexcel.config.HeaderConfig;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.hssf.usermodel.HSSFRichTextString;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.lang.reflect.Constructor;
import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.text.DateFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.Map;

/**
 * excel文件解析
 *
 * @author jiaxuch
 * @data 2020/6/26
 */
public class ExcelParseUtils {

    public static Logger logger = LoggerFactory.getLogger(ExcelParseUtils.class);

    public static final String XLSX = "xlsx";

    /**
     * 使用默认日期格式（yyyy-MM-dd）
     *
     * @param file
     * @param map
     * @param clazz
     * @param <T>
     * @return
     * @throws Exception
     */
    public static <T> List<T> parseExcelToList(MultipartFile file, Map<String, String> map, boolean flag, SaveDataInterface saveDataInterface, Class<T> clazz) throws Exception {
       return parseExcelToList(file, map, null, flag, saveDataInterface, clazz);
    }

    /**
     * 提供SaveDataInterface默认实现，不保存任何数据，并使用默认日期格式化（yyyy-MM-dd）
     *
     * @param file
     * @param map
     * @param flag
     * @param clazz
     * @param <T>
     * @return
     * @throws Exception
     */
    public static <T> List<T> parseExcelToList(MultipartFile file, Map<String, String> map, boolean flag, Class<T> clazz) throws Exception {
        //默认不存数据
        SaveDataInterface saveDataInterface = new SaveDataInterface() {
            @Override
            public int save(Object o) {
              return 0;
            }
        };
        return parseExcelToList(file, map, null, flag, saveDataInterface, clazz);
    }

    /**
     * 将上传的文件解析成集合对象
     * 该方法只支持简单数据boolean, String, int, double,long,java.util.Date
     * 如果excel获取的类型和对象不一致，则认为对象中类型为字符串，将获取的值全部转换成String类型进行赋值。
     *
     * @param file 传文件流
     * @param map excel中列对应实体类中的属性，以0开头
     * @param dateFormat 如果有时间类型，时间类型格式化成什么样的字符串
     * @param flag 标识excle表格第一行是否需要解析成对象。默认不需要
     * @param saveDataInterface 钩子对象，如果需要保存数据，请实现接口，并重写方法（主要避免返回的list重新遍历存值）
     * @param clazz 将数据封装的目标对象类型
     * @param <T> 实体类的类型
     * @return
     * @throws Exception
     */
    public static <T> List<T> parseExcelToList(MultipartFile file, Map<String, String> map, DateFormat dateFormat, boolean flag, SaveDataInterface saveDataInterface, Class<T> clazz) throws Exception {
        //文件扩展名校验及是否是office文件校验
        checkExtensionAndType(file);

        List<T> list = new ArrayList<T>();
        try {
            //正确的文件类型 自动判断2003或者2007
            Workbook workbook = PoiUtils.getWorkbookAuto(file);

            //默认只有一个sheet
            Sheet sheet = workbook.getSheetAt(0);

            //获得sheet有多少行
            int rows = sheet.getPhysicalNumberOfRows();

            //读第一个sheet
            for (int i = 0; i < rows; i++){

                //flag为ture标识第一行需要解析，false为不需要解析
                if(!flag && i == 0){
                    continue;
                }
                Row row = sheet.getRow(i);
                //通过反射获取无产构造器
                Constructor<T> constructor = clazz.getConstructor();
                //创建对应的对象
                T t = constructor.newInstance();

                //遍历每行中的单元
                for (int j = 0; j < row.getLastCellNum(); j++){

                    // 列，每一列在map中对应着一个对象的属性
                    Cell cell = row.getCell(j);

                    if (cell != null){
                        //处理单元格数据
                        parseExcelCell(map, dateFormat, clazz, t, j, cell);
                    }
                }
                // 是不是可以放一个钩子，直接让别人创建接口子类，然后调用保存方法。
                saveDataInterface.save(t);
                list.add(t);
            }

        } catch (IOException e) {
            e.printStackTrace();
        }

        return list;
    }

    /**
     * 处理单元格
     *
     * @param map 单元格列号和将要封装的对象属性对应映射
     * @param dateFormat
     * @param clazz
     * @param t
     * @param j
     * @param cell
     * @param <T>
     * @throws InvocationTargetException
     * @throws IllegalAccessException
     * @throws NoSuchMethodException
     * @throws NoSuchFieldException
     */
    private static <T> void parseExcelCell(Map<String, String> map, DateFormat dateFormat, Class<T> clazz, T t, int j, Cell cell) throws InvocationTargetException, IllegalAccessException, NoSuchMethodException, NoSuchFieldException, ParseException {
        // 获取属性名称
        String fieldName = map.get(j + "");
        // 根据属性名获取属性对象
        Field field = clazz.getDeclaredField(fieldName);
        // 获取属性类型
        Class<?> fieldType = field.getType();

        // 获取属性对应的set方法名称
        String methodName = getFieldSetOrGetMethod(fieldName, "set");

        // 获取set方法对象
        Method method = clazz.getMethod(methodName, fieldType);
        logger.info("获取方法名：" + method.getName());
        // 默认时间格式化类型
        if(dateFormat == null){
            dateFormat = new SimpleDateFormat("yyyy-MM-dd");
        }
        // 获取单元格属性
        CellType cellTypeEnum = cell.getCellTypeEnum();
        // 布尔类型
        if(CellType.BOOLEAN.toString().equals(cellTypeEnum.toString())){
            handleBoolean(t, cell, fieldType, method);
            // 数值类型
        }else if(CellType.NUMERIC.toString().equals(cellTypeEnum.toString())){
            // 时间类型处理
            if (HSSFDateUtil.isCellDateFormatted(cell)) {
                handleDataNumeric(dateFormat, t, cell, fieldType, method);
            // 其他数值类型
            }else {
                handleOutherNumeric(t, cell, fieldType, method);
            }
        // 字符串
        }else if(CellType.STRING.toString().equals(cellTypeEnum.toString())){
            //上传日期是字符串，但是属性是时间类型
            if(fieldType == Date.class){
                handelDateStr(dateFormat, t, cell, method);
            }else {
                handleString(t, cell, method);
            }
        }
    }

    private static <T> void handelDateStr(DateFormat dateFormat, T t, Cell cell, Method method) throws ParseException, IllegalAccessException, InvocationTargetException {
        String datestr = cell.getStringCellValue();
        Date parse = dateFormat.parse(datestr);
        method.invoke(t, parse);
    }

    /**
     * 字符串类型处理
     *
     * @param t
     * @param cell
     * @param method
     * @param <T>
     * @throws InvocationTargetException
     * @throws IllegalAccessException
     */
    private static <T> void handleString(T t, Cell cell, Method method) throws InvocationTargetException, IllegalAccessException {
        String stringCellValue = cell.getStringCellValue();
        System.out.println(stringCellValue);
        method.invoke(t, stringCellValue);
    }

    /**
     * boolean 所有boolean类型全部转成字符串类型
     *
     * @param t
     * @param cell
     * @param fieldType
     * @param method
     * @param <T>
     * @throws InvocationTargetException
     * @throws IllegalAccessException
     */
    private static <T> void handleBoolean(T t, Cell cell, Class<?> fieldType, Method method) throws InvocationTargetException, IllegalAccessException {
        cell.setCellType(CellType.STRING);
        String stringCellValue = cell.getStringCellValue();
        System.out.println(stringCellValue);
        //同一类型，直接赋值，否则，转成字符串赋值
        method.invoke(t, stringCellValue);

    }

    /**
     * 时间类型处理
     *
     * @param dateFormat
     * @param t
     * @param cell
     * @param fieldType
     * @param method
     * @param <T>
     * @return
     * @throws IllegalAccessException
     * @throws InvocationTargetException
     */
    private static <T> DateFormat handleDataNumeric(DateFormat dateFormat, T t, Cell cell, Class<?> fieldType, Method method) throws InvocationTargetException, IllegalAccessException {
        Date data = cell.getDateCellValue();
        if(compareObjectAndCellType(Date.class, fieldType)){
            method.invoke(t, data);
        }else {
            String format = dateFormat.format(data);
            System.out.println(format);
            method.invoke(t, format);
        }
        return dateFormat;
    }

    /**
     * 其他数值类型
     *
     * @param t
     * @param cell
     * @param fieldType
     * @param method
     * @param <T>
     * @throws IllegalAccessException
     * @throws InvocationTargetException
     */
    private static <T> void handleOutherNumeric(T t, Cell cell, Class<?> fieldType, Method method) throws InvocationTargetException, IllegalAccessException {
//        cell.setCellType(CellType.STRING);
//        String cellValue = cell.getStringCellValue();
        double cellValue = cell.getNumericCellValue();
        System.out.println(cellValue);
        if(compareObjectAndCellType(double.class, fieldType) || compareObjectAndCellType(Double.class, fieldType)){
            method.invoke(t, cellValue);
        }else if(compareObjectAndCellType(int.class, fieldType) || compareObjectAndCellType(Integer.class, fieldType)){
            method.invoke(t, new Double(cellValue).intValue());
        }else if(compareObjectAndCellType(long.class, fieldType) || compareObjectAndCellType(Long.class, fieldType)){
            method.invoke(t, new Double(cellValue).longValue());
        }else {
            method.invoke(t, cellValue + "");
        }
    }

    /**
     * 对输入文件扩展名及类型做校验
     *
     * @param file
     * @throws Exception
     */
    private static void checkExtensionAndType(MultipartFile file) throws Exception {
        if (!PoiUtils.checkExtension(file)){
            new Exception("请求文件类型错误:后缀名错误");
        }
        if (!PoiUtils.isOfficeFile(file)) {
            new Exception("请求文件类型错误:文件类型错误");
        }
    }

    /**
     * 比较两个Class是否是同一个
     *
     * @param booleanClass
     * @param fieldType
     * @return
     */
    private static boolean compareObjectAndCellType(Class<?> booleanClass, Class<?> fieldType) {
        return booleanClass == fieldType;
    }

    /**
     * 根据字段名获取对应的set方法名称
     *
     * @param field
     * @return
     */
    public static String getFieldSetOrGetMethod(String field, String prefix) {
        //将获取的第一个字符串转成大写
        String first = field.substring(0, 1).toUpperCase();
        //substring(1),获取索引位置1后面所有剩余的字符串
        String after = field.substring(1);
        return prefix + first + after;
    }



    /**
     *
     * @param list
     * @param title
     * @param fileName xls,xlsx
     * @param response
     * @param
     */
    public static <T> void parseListToExcel(List<T> list, Map<String, HeaderConfig> title, String fileName, HttpServletResponse response) throws Exception {
        Workbook workbook = getSheets(fileName);
        Sheet sheet = workbook.createSheet("信息表");
        int rowNum = 0;
        int size = title.size();
        //表格头数据不为空
        if(title != null && title.size() != 0){
            Row row = sheet.createRow(0);
            rowNum++;
            for (int i = 0; i < size; i++){
                HeaderConfig headerConfig = title.get(i + "");
                //headers表示excel表中第一行的表头
                Cell cell = row.createCell(i);
                //根据文件名创建RichTextString
                RichTextString text = getRichTextString(fileName, headerConfig);
                cell.setCellValue(text);
            }
        }

        //新增数据行，并且设置单元格数据
        //在表中存放查询到的数据放入对应的列
        for (T t : list) {
            Row row1 = sheet.createRow(rowNum);
            //获取Class对象
            Class<?> tClass = t.getClass();

            //给行中的每一个数据进行赋值
            for (int i = 0; i < size; i++){
                //获取方法名
                String fieldName = title.get(i + "").getFieldName();
                String getFileName = ExcelParseUtils.getFieldSetOrGetMethod(fieldName, "get");
                //获取对象具体方法
                Method method = tClass.getMethod(getFileName);
                Object invoke = method.invoke(t);

                Cell cell = row1.createCell(i);
                createCell(cell, invoke);

            }
            rowNum++;
        }
        //输出文件流
        outputExcel(fileName, response, workbook);
    }

    /**
     * 根据文件名判断创建什么样的Workbook （XSSFWorkbook or HSSFWorkbook）
     *
     * @param fileName
     * @return
     */
    private static Workbook getSheets(String fileName) {
        Workbook workbook;
        if(fileName.endsWith(XLSX)){
            workbook = new XSSFWorkbook();
        } else{
            workbook = new HSSFWorkbook();
        }
        return workbook;
    }

    /**
     * 根据文件后缀判断创建什么类型的RichTextString （XSSFRichTextString or HSSFRichTextString）
     *
     * @param fileName
     * @param headerConfig
     * @return
     */
    private static RichTextString getRichTextString(String fileName, HeaderConfig headerConfig) {
        RichTextString text;
        if(fileName.endsWith(XLSX)){
            text = new XSSFRichTextString(headerConfig.getHeaderName());
        }else {
            text = new HSSFRichTextString(headerConfig.getHeaderName());

        }
        return text;
    }

    /**
     * 将excel通过输入流输入给前端
     *
     * @param fileName
     * @param response
     * @param workbook
     */
    private static void outputExcel(String fileName, HttpServletResponse response, Workbook workbook) {
        response.setContentType("application/octet-stream");
        response.setHeader("Content-disposition", "attachment;filename=" + fileName);
        try {
            response.flushBuffer();
            workbook.write(response.getOutputStream());
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    /**
     * Excel单元格设置
     *
     * @param cell
     * @param invoke
     */
    private static void createCell(Cell cell, Object invoke) {
        if(invoke.getClass() == double.class || invoke.getClass() == Double.class){
            cell.setCellValue((double)invoke);
        }else if (invoke.getClass() == long.class || invoke.getClass() == Long.class){
            cell.setCellValue((long)invoke);
        }else if (invoke.getClass() == int.class || invoke.getClass() == Integer.class){
            cell.setCellValue((int)invoke);
        }else if (invoke.getClass() == Date.class){
            SimpleDateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd");
            String format = dateFormat.format(invoke);
            System.out.println(format);
            cell.setCellValue(format);
//            cell.setCellValue((Date) invoke);
        }else {
            cell.setCellValue((String) invoke);
        }
    }

}
