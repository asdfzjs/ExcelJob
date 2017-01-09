package com.zillionfortune.main;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import javax.servlet.http.HttpServletResponse;

import org.apache.poi.openxml4j.util.ZipSecureFile.ThresholdInputStream;
import org.apache.poi.ss.usermodel.Workbook;

import com.zillionfortune.serviceImpl.ExcelOperImpl;
import com.zillionfortune.util.DBHelper;

public class main {
	static String sql1 = null; 
	static String sql2 = null;
    static DBHelper db1 = null;    
    static ResultSet ret = null;
	public static void main(String[] args) {
		//String date = args[0];// 查询的时间
		sql1 = " SELECT modifyUser, '登录' , modifyDate   FROM ssp.operateLog where logSource = 888";//SQL语句    
		db1 = new DBHelper(sql1);//创建DBHelper对象    
    
        try{    
            ret = db1.pst.executeQuery();//执行语句，得到结果集   
            List<Map<String, Object>> userlist = new ArrayList<Map<String,Object>>();
            while (ret.next()) {    
                String modifyUser = ret.getString(1);    
                String action = ret.getString(2);    
                String modifyDate = ret.getTimestamp(3).toString();    
                Map<String,Object> a = new HashMap<String,Object>();
                a.put("modifyUser", modifyUser);
                a.put("action", action);
                a.put("modifyDate", modifyDate);
                userlist.add(a);
            }//显示数据    
            ret.close();    
            db1.close();//关闭连接 
            db1.exportPersonalExcel("2016-11-12",userlist);
        } catch (SQLException e) {    
            e.printStackTrace();    
        }    
	}
	
	
}
