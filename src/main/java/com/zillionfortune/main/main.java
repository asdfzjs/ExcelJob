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
		//String date = args[0];// ��ѯ��ʱ��
		sql1 = " SELECT modifyUser, '��¼' , modifyDate   FROM ssp.operateLog where logSource = 888";//SQL���    
		db1 = new DBHelper(sql1);//����DBHelper����    
    
        try{    
            ret = db1.pst.executeQuery();//ִ����䣬�õ������   
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
            }//��ʾ����    
            ret.close();    
            db1.close();//�ر����� 
            db1.exportPersonalExcel("2016-11-12",userlist);
        } catch (SQLException e) {    
            e.printStackTrace();    
        }    
	}
	
	
}
