package com.zillionfortune.util;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.sql.DriverManager;
import java.sql.SQLException;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.Workbook;

import com.mysql.jdbc.Connection;
import com.mysql.jdbc.PreparedStatement;
import com.zillionfortune.serviceImpl.ExcelOperImpl;

public class DBHelper {    
    public static final String url = "jdbc:mysql://ip/ssp";    
    public static final String name = "com.mysql.jdbc.Driver";    
    public static final String user = "user";    
    public static final String password = "pwd";    
    
    public java.sql.Connection conn = null;    
    public PreparedStatement pst = null;    
    
    public DBHelper(String sql) {    
        try {    
            Class.forName(name);//指定连接类型    
            conn = DriverManager.getConnection(url, user, password);//获取连接    
            pst = (PreparedStatement) conn.prepareStatement(sql);//准备执行语句    
        } catch (Exception e) {    
            e.printStackTrace();    
        }    
    }    
    
    public void close() {    
        try {    
            this.conn.close();    
            this.pst.close();    
        } catch (SQLException e) {    
            e.printStackTrace();    
        }    
    } 
    
public  void exportPersonalExcel(String date,List<Map<String, Object>> userlist) {
		
		ExcelOperImpl a = new ExcelOperImpl();
		String fileDir = "c:";
		File dir = new File(fileDir);
		if(!dir.exists()){
			dir.mkdirs();
		}
		String fileName  = "日志"+date+".xlsx";
		
		String filePath = fileDir+"/"+fileName;
		File file = new File(filePath);
		
		System.out.println(file.getAbsolutePath());
		
		OutputStream out = null;
		FileOutputStream fileOutputStream = null;
		FileInputStream in = null;
		try{
			//如果物理空间不存在，则存储一份
			if(!file.exists()){
				fileOutputStream = new FileOutputStream(new File(filePath));
				Workbook workbook =  a.exportToExcelWithTemplet("a.xlsx", userlist);
				workbook.write(fileOutputStream);
				fileOutputStream.flush();
			}
		} catch (IOException e) {
			//删除生成的物理文件
			File file2 = new File(filePath);
			file2.delete();
			e.printStackTrace();
			throw new RuntimeException("报表导出失败:"+e.getMessage());
		}finally {
			if(null != in){
				try {
					in.close();
				} catch (IOException e) {
					e.printStackTrace();
				}
			}
			if(null != out){
				try {
					out.close();
				} catch (IOException e) {
					e.printStackTrace();
				}
			}
			if(null != fileOutputStream){
				try {
					fileOutputStream.close();
				} catch (IOException e) {
					e.printStackTrace();
				}
			}
		}
		
	}
}    
