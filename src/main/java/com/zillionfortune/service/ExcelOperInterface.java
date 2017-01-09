package com.zillionfortune.service;

import org.apache.poi.ss.usermodel.Workbook;

import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.List;

import javax.servlet.http.HttpServletResponse;

/**
 * Created by zhangwenjun on 2016/11/14.
 */
public interface ExcelOperInterface {
    <T> List<T> readFromExcel(Class<T> entityClass) throws IOException, IllegalAccessException, InstantiationException;
    void init(String fileName, InputStream excelInfo) throws IOException;
    Workbook exportToExcelWithTemplet(String templetName, List... bean) throws IOException;
    void exportToExcelWithTemplet4Web(String templetName, HttpServletResponse response, String downloadFileName, List... bean) throws IOException;

}
