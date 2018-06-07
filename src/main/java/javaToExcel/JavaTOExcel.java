package javaToExcel;

import java.io.FileOutputStream;
import java.sql.Connection;
import java.sql.DatabaseMetaData;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.ResultSetMetaData;
import java.sql.Statement;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import com.alibaba.fastjson.JSONObject;

public class JavaTOExcel {
	
	//@Test
	public void createXls() throws Exception{
	  //声明一个工作薄
	  HSSFWorkbook wb = new HSSFWorkbook();
	  //声明表
	  HSSFSheet sheet = wb.createSheet("第一个表");
	  //声明行
	  HSSFRow row = sheet.createRow(7);
	  //声明列
	  HSSFCell cel = row.createCell(3);
	  //写入数据
	  cel.setCellValue("你也好");
	   
	  FileOutputStream fileOut = new FileOutputStream("d:/a/b.xls");
	  wb.write(fileOut);
	  fileOut.close();
	 }
	
	
	
	public static void export() throws Exception{
		
		String url="jdbc:mysql://172.19.56.127:3306/cfh_activity?useUnicode=true&characterEncoding=utf8";
        String user="cfh_activity";
        String password="Cfh_activity123";
        Connection con=DriverManager.getConnection(url, user, password);//获得连接
        System.err.println(con.toString());
		  //声明需要导出的数据库
		  String dbName = "cfh_activity";
		  //声明book
		  HSSFWorkbook book = new HSSFWorkbook();
		  //获取Connection,获取db的元数据
		  //Connection con = DataSourceUtils.getConn();
		  //声明statemen
		  Statement st = con.createStatement();
		  //st.execute("use "+dbName);
		  DatabaseMetaData dmd = con.getMetaData();
		  System.err.println("dmd:"+dmd.getUserName());
		  //获取数据库有多少表
		  ResultSet rs = dmd.getTables(dbName,dbName,"",new String[]{"TABLE"});
		  //获取所有表名　－　就是一个sheet
		  List<String> tables = new ArrayList<String>();
		 
		  while(rs.next()){
		   String tableName = rs.getString("TABLE_NAME");
		   tables.add(tableName);
		  }
		  System.err.println(tables.toString());
		  for(String tableName:tables){
		   HSSFSheet sheet = book.createSheet(tableName);
		   //声明sql
		   String sql = "SELECT COLUMN_NAME 列名, COLUMN_TYPE 数据类型,DATA_TYPE 字段类型,CHARACTER_MAXIMUM_LENGTH 长度,IS_NULLABLE 是否可为空,COLUMN_DEFAULT 默认值,COLUMN_COMMENT 备注 FROM INFORMATION_SCHEMA.COLUMNS where table_schema ='"
		   		+ dbName + "' AND table_name  = '" + tableName + "'"; 
		   // "select * from "+dbName+"."+tableName;
		   //System.err.println("sql:"+sql);、】
		   String[] clnStr = new String[]{"列名","数据类型","字段类型","长度","是否可为空","默认值","备注"};
		   //查询数据		   
		   rs = st.executeQuery(sql);
		   //根据查询的结果，分析结果集的元数据
		   ResultSetMetaData rsmd = rs.getMetaData();
		   //获取这个查询有多少列
		   int cols = rsmd.getColumnCount();
		   //创建第三行
		   HSSFRow row = sheet.createRow(2);
		   for(int i=0;i<cols;i++){			   
			    String colName = clnStr[i];
			    //创建一个新的列
			    HSSFCell cell = row.createCell(i);
			    //写入列名
			    cell.setCellValue(colName);
		   }
		   //获取所有列名
		   /*for(int i=0;i<cols;i++){
		    String colName = rsmd.getColumnName(i+1);
		    //创建一个新的列
		    HSSFCell cell = row.createCell(i);
		    //写入列名
		    cell.setCellValue(colName);
		   }*/
		   //遍历数据
		   int index = 3;
		   while(rs.next()){
		    row = sheet.createRow(index++);
		    //声明列
		    for(int i=0;i<cols;i++){
		     String val = rs.getString(i+1);
		     //声明列
		     HSSFCell cel = row.createCell(i);
		     //放数据
		     cel.setCellValue(val);
		    }
		   }
		  }
		  con.close();
		  book.write(new FileOutputStream("C:/Users/x_lilong/Desktop/数据库表/"+dbName+".xls"));
		 }

	public static void main(String[] args) {
		try {
			export();
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}

}
