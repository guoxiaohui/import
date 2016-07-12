package com.javen.db;

import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;
/**
 * 数据库工具类
 * @author guoxiaoHui
 * @since 2016-7-6 13:18:06
 */
public class DBhepler {
	
	private String driver = "com.mysql.jdbc.Driver";
	private String url = "jdbc:mysql://localhost/test";
    private String userName = "root";
    private String passWord = "p@ssw0rd";
    
    private Connection con = null;
    
    public Connection getConnection() {
        try {
            Class.forName(driver);
            con = DriverManager.getConnection(url, userName, passWord);
            System.out.println("连接数据库成功");
        } catch (ClassNotFoundException e) {
            // TODO Auto-generated catch block
              System.err.println("装载 JDBC/ODBC 驱动程序失败。" );  
            e.printStackTrace();
        } catch (SQLException e) {
            // TODO Auto-generated catch block
            System.err.println("无法连接数据库" ); 
            e.printStackTrace();
        }
        
        return con;
    }
    
    /**
	 * 关闭连接
	 */
	public void closeConnection(Connection connn,Statement stm,ResultSet set){
		try{
			if(set != null){
				set.close();
			}
			
			if(stm != null){
				stm.close();
			}
			
			if(con != null){
				con.close();
			}
		}catch(Exception e){
			e.printStackTrace();
		}
	}
    
}
