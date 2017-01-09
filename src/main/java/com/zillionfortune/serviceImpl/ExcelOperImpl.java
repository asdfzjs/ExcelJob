package com.zillionfortune.serviceImpl;


import java.io.BufferedOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.Set;

import javax.servlet.http.HttpServletResponse;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.zillionfortune.service.ExcelOperInterface;
import com.zillionfortune.util.CellToField;
import com.zillionfortune.util.ExcelFormatType;
import com.zillionfortune.util.ExportExcelException;
import com.zillionfortune.util.SheetToClass;


/**
 * Created by zhangwenjun on 2016/11/14.
 */

public class ExcelOperImpl implements ExcelOperInterface {

    private Workbook workbook;

    public enum ExcelVersion{
        OFFICE_EXCEL_2003_FILEFIX("xls"),
        OFFICE_EXCEL_2010_FILEFIX("xlsx");

        String version;

        private ExcelVersion(String version){
            this.version = version;
        }

        @Override
        public String toString() {
            return this.version;
        }
    }

    public ExcelOperImpl(){

    }

    public void init(String fileName,InputStream excelInfo) throws IOException {
        String version = "";

        if (fileName == null || fileName.trim().isEmpty()) {
            throw new ExportExcelException("excel名称不能为空");
        }

        if (fileName.contains(".")) {
            version = fileName.substring(fileName.lastIndexOf(".") + 1, fileName.length());
        }

        if(version.equals(ExcelVersion.OFFICE_EXCEL_2003_FILEFIX.toString())){
            workbook = new HSSFWorkbook(excelInfo);
        }else if (version.equals(ExcelVersion.OFFICE_EXCEL_2010_FILEFIX.toString())){
            workbook = new XSSFWorkbook(excelInfo);
        }
    }

    public <T> List<T> readFromExcel(Class<T> entityClass) throws IOException, IllegalAccessException, InstantiationException {
        //公司—0	工号-1	姓名-2	部门-3	分公司-4	业务部-5	城市-6	大区负责人-7 分公司负责人-8	业务总监-9	理财经理-10
        // 	岗位-11	岗位类别-12	职级-13	入职日期-14
        Sheet sheet = workbook.getSheet(entityClass.getAnnotation(SheetToClass.class).sheetName());
        if(sheet == null){
            throw new ExportExcelException("导入EXCEL 错误,需要有 <"+entityClass.getAnnotation(SheetToClass.class).sheetName()+"> sheet,请下载模板进行修改");
        }
        List<T> result = new ArrayList<T>();
        for (int rowNum = 1; rowNum <= sheet.getLastRowNum(); rowNum++) {
            Row row = sheet.getRow(rowNum);
            if(row == null) {
                continue;
            }
            CellToField ctf = null;
            try {
                T empty =  (T)entityClass.newInstance();
                Field[] fields = entityClass.getDeclaredFields();

                for(Field field : fields){
                    ctf = field.getAnnotation(CellToField.class);
                    if(ctf != null){
                        String value = "";
                        switch(ctf.format()){
                            case STRING:
                                value = row.getCell(ctf.cellIndex()) == null? "":row.getCell(ctf.cellIndex()).getStringCellValue().trim();
                                break;
                            case DATE:
                                SimpleDateFormat format = new SimpleDateFormat("yyyy-MM-dd");
                                try {
                                    value = format.format((row.getCell(ctf.cellIndex()).getDateCellValue()));
                                } catch (Exception e) {
                                    throw new ExportExcelException("<" + entityClass.getAnnotation(SheetToClass.class).sheetName() +  "> sheet, 第"+(rowNum+1)+"行,第"+(ctf.cellIndex()+1)+"列数据请用YYYY-MM-DD的文本格式填写");
                                }
                                break;
                            case STRINGDATE:
                                SimpleDateFormat stringFormat = new SimpleDateFormat("yyyy-MM-dd");
                                try {
                                    value = stringFormat.format(stringFormat.parse(row.getCell(ctf.cellIndex()).getStringCellValue()));
                                } catch (Exception e) {
                                    throw new ExportExcelException("<" + entityClass.getAnnotation(SheetToClass.class).sheetName() +  "> sheet, 第"+(rowNum+1)+"行,第"+(ctf.cellIndex()+1)+"列数据请用YYYY-MM-DD的文本格式填写");
                                }
                                break;
                            case INT:
                                value = Double.valueOf(row.getCell(ctf.cellIndex()).getNumericCellValue()).intValue()+"";
                                break;
                            case DOUBLE:
                                value = row.getCell(ctf.cellIndex()).getNumericCellValue()+"";
                                break;
                            default:
                                value = getXlsValue(row.getCell(ctf.cellIndex()));
                        }
                        field.setAccessible(true);
                        if(ctf.notNull() && StringUtils.isEmpty(value)){
                            throw new ExportExcelException("<" + entityClass.getAnnotation(SheetToClass.class).sheetName() +  "> sheet, 第"+(rowNum+1)+"行,第"+(ctf.cellIndex()+1)+"列数据不能为空");
                        }else{
                            field.set(empty,value);
                        }

                    }
                }
                result.add(empty);

            } catch (IllegalStateException e) {
                e.printStackTrace();
                if(ctf.format() == ExcelFormatType.STRINGDATE || ctf.format() == ExcelFormatType.DATE){
                    throw new ExportExcelException("<"+entityClass.getAnnotation(SheetToClass.class).sheetName()+"> sheet, 第"+(rowNum+1)+"行,第"+(ctf.cellIndex()+1)+"列数据格式不对,请用文本输入YYYY-MM-DD");
                }else{
                    throw new ExportExcelException("<"+entityClass.getAnnotation(SheetToClass.class).sheetName()+"> sheet, 第"+(rowNum+1)+"行,第"+(ctf.cellIndex()+1)+"列数据格式不对,请用文本输入");
                }

            } catch (IllegalAccessException e) {
                e.printStackTrace();
                throw e;
            } catch (InstantiationException e) {
                e.printStackTrace();
                throw e;
            }
        }

        return result;
    }

    private String getXlsValue(Cell cell) {
        if(cell == null){
            return  "";
        }
        if (cell.getCellType() == cell.CELL_TYPE_BOOLEAN) {
            return String.valueOf(cell.getBooleanCellValue());
        } else if (cell.getCellType() == cell.CELL_TYPE_NUMERIC) {
            return String.valueOf(cell.getNumericCellValue());
        } else {
            return String.valueOf(cell.getStringCellValue());
        }
    }

    public Workbook exportToExcelWithTemplet(String templetName,List... bean) throws IOException {
        //对excel进行初始化并且封装对象
        try {
            init(templetName,getClass().getResourceAsStream("/template/"+templetName));
        } catch (IOException e) {
            e.printStackTrace();
        }

        List[] sheetValue = (List[])bean;

        for(int i = 0 ; i < sheetValue.length ; i ++) {
            Sheet sheet = workbook.getSheetAt(i);
            int startRow = 0;
            //记录开始行，记录对应的字段
            Map<Integer, String> mappedField = new HashMap<>();
            for (int rowNum = 0; rowNum <= sheet.getLastRowNum(); rowNum++) {
                Row row = sheet.getRow(rowNum);
                if(row == null){
                	continue;
                }
                short firstCellNum = row.getFirstCellNum();
                short lastCellNum = row.getLastCellNum();
                for (; firstCellNum < lastCellNum; firstCellNum++) {
                    Cell cell = row.getCell(firstCellNum);
                    if(cell == null){
                        continue;
                    }
                    String x = cell.getStringCellValue();
                    //记录约定的字符，并且
                    if (x != null && x.startsWith("${") && x.endsWith("}")) {
                        mappedField.put(Integer.parseInt(firstCellNum + ""), x.substring(2, x.length() - 1));
                        startRow = rowNum;
                    } else {
                        continue;
                    }
                }
            }

            //清除第一行
            sheet.createRow(startRow);

            if (sheetValue[i] == null) {
                return workbook;
            }


            for (Object beanInfo : sheetValue[i]) {
                try {
                    setRowValue(mappedField, sheet.createRow(startRow), beanInfo);
                    startRow++;
                } catch (NoSuchMethodException e) {
                    e.printStackTrace();
                } catch (InvocationTargetException e) {
                    e.printStackTrace();
                } catch (IllegalAccessException e) {
                    e.printStackTrace();
                }
            }

        }
        workbook.setActiveSheet(0);
        return  workbook;
    }

    private void setRowValue(Map<Integer,String> mappedField, Row row,Object valueBean) throws NoSuchMethodException, InvocationTargetException, IllegalAccessException {
        Set<Integer> keySet = mappedField.keySet();
        Iterator<Integer>  iterator = keySet.iterator();
        while(iterator.hasNext()){
            Integer cellIndex = iterator.next();
            Cell cell = row.createCell(cellIndex);
            CellStyle cellStyle = workbook.createCellStyle();
            DataFormat format = workbook.createDataFormat();
            cellStyle.setDataFormat(format.getFormat("@"));
            cell.setCellStyle(cellStyle);
            //判断具体的类型
            if(valueBean instanceof Map){
                //通过get获得内容
                Object value = ((Map)valueBean).get(mappedField.get(cellIndex));
                if(value == null){
                    value = "";
                }
                cell.setCellValue(value.toString());
            }else{
                //通过反射获得内容
                Class beanClass = valueBean.getClass();
                Method getMethod = beanClass.getMethod("get"+captureName(mappedField.get(cellIndex)));
                Object value = getMethod.invoke(valueBean);
                if(value == null){
                    value = "";
                }
                cell.setCellValue(value.toString());
            }
        }

    }

    private String captureName(String name) {
        char[] cs=name.toCharArray();
        cs[0]-=32;
        return String.valueOf(cs);

    }

    public void exportToExcelWithTemplet4Web(String templetName, HttpServletResponse response, String downloadFileName, List... bean) throws IOException {
        response.setContentType("application/x-excel");
        response.setCharacterEncoding("UTF-8");
        response.setHeader("Content-Disposition", "attachment; filename="+java.net.URLEncoder.encode(downloadFileName, "UTF-8"));
        BufferedOutputStream out = new BufferedOutputStream(response.getOutputStream());
        exportToExcelWithTemplet(templetName,bean).write(out);
    }
}

