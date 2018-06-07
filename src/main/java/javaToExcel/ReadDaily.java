package javaToExcel;

import java.io.File;
import java.util.List;

public class ReadDaily {
	
	@SuppressWarnings({ "rawtypes", "static-access", "unused" })
	public static void main(String[] args) {

        ReadExcel obj = new ReadExcel();
        // 此处为我创建Excel路径：E:/zhanhj/studysrc/jxl下
        File file = new File("D://0_1//蚂蚁财富号//01_每日财经资料//4yue 20haohou//01麻辣财经-201804.xls");
        List excelList = obj.readExcel(file);
        String sql = "INSERT INTO daily_finance_info (date,title,content,comment,self_comment,hot_plate)VALUES(";
        String sqlDelete = "DELETE FROM daily_finance_info WHERE date in (";
        String sqlSelect = "SELECT * from daily_finance_info WHERE date in (";
        String dateString = "";
        String sqlend = ");";
        System.out.println("list中的数据打印出来");
        for (int i = 1; i < excelList.size(); i++) {
        	List list = (List) excelList.get(i);
        	if(i == 1){
        		dateString = dateString.concat("'"+list.get(0).toString()+"'");
        	}else{
        		dateString = dateString.concat(",'"+list.get(0).toString()+"'");
        	}
        }
        System.err.println(sqlDelete.concat(dateString).concat(sqlend));
        for (int i = 1; i < excelList.size(); i++) {
            List list = (List) excelList.get(i);
            StringBuffer sqlValue = new StringBuffer();
        	StringBuffer opptions = new StringBuffer();
            for (int j = 0; j < list.size(); j++) {
            	if(j == 0){
            		sqlValue.append("'"+list.get(j)+"'");
            		
            	}else if(" ".equals(list.get(j).toString())){
            		sqlValue.append(",null");
            	}else {
            		sqlValue.append(",'"+list.get(j)+"'");
            		//System.out.print(list.get(j));
            	}
                
            }
            System.err.println(sql.concat(sqlValue.toString()).concat(sqlend));
        }
        
        System.err.println(sqlSelect.concat(dateString).concat(sqlend));

    
	}

}
