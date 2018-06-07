package javaToExcel;

import java.io.BufferedWriter;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.io.Writer;
import java.util.ArrayList;
import java.util.List;

import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;
public class ReadWorldCupExcel {
    public static void main(String[] args) {
        ReadWorldCupExcel obj = new ReadWorldCupExcel();
        // 此处为我创建Excel路径：E:/zhanhj/studysrc/jxl下
        //File file = new File("D:\\0_1\\蚂蚁财富号\\20180426_每日财经和问答红包\\问答红包\\题库2018.4.9-筛选.xls");
        File file = new File("D:\\0_1\\蚂蚁财富号\\20180531_世界杯竞猜\\20180606.xls");
        List excelList = obj.readExcel(file);
        String sql = "INSERT INTO world_cup_competition_info (id,act_type,stage,start_time,last_time,end_time,match_stage,match_teamA,match_teamB,url_A,url_B) VALUES (";
        String sqlend = ");";
        String nextLineStr = "\r\n";
        String deletSql = "TRUNCATE TABLE world_cup_competition_info;";
        StringBuffer insertSql = new StringBuffer();
        insertSql.append(deletSql).append(nextLineStr);
        System.out.println("list中的数据打印出来");
        for (int i = 1; i < excelList.size(); i++) {
        	
            List list = (List) excelList.get(i);
            StringBuffer sqlValue = new StringBuffer();
            for (int j = 0; j < list.size(); j++) {
            	String value = list.get(j).toString();
            	if(j == 0){
            		sqlValue.append(value+",");
            	}else if(null == value || "".equals(value) || " ".equals(value)){
            		sqlValue.append(""+null+",");
            	}else{
            		sqlValue.append("'"+value+"',");
            	}
            }
            String subSql = sql.concat(sqlValue.toString());
            subSql = subSql.substring(0, subSql.length()-1);
            System.err.println(subSql.concat(sqlend));
            insertSql.append(subSql.concat(sqlend)).append(nextLineStr);
        }
        
        String selSql = "SELECT * FROM world_cup_competition_info;";
        insertSql.append(selSql);
        System.err.println("=====================================");
        System.err.println(insertSql.toString());
        try {
			OutputStream fos=new FileOutputStream("D:\\0_1\\蚂蚁财富号\\20180531_世界杯竞猜\\sql\\20180601_dml_world_cup_competition_info.sql");
			fos.write(insertSql.toString().getBytes());
			fos.flush();
			fos.close();
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}  

    }
    // 去读Excel的方法readExcel，该方法的入口参数为一个File对象
    public static List readExcel(File file) {
        try {
            // 创建输入流，读取Excel
            InputStream is = new FileInputStream(file.getAbsolutePath());
            // jxl提供的Workbook类
            Workbook wb = Workbook.getWorkbook(is);
            // Excel的页签数量
            int sheet_size = wb.getNumberOfSheets();
            for (int index = 0; index < sheet_size; index++) {
                List<List> outerList=new ArrayList<List>();
                // 每个页签创建一个Sheet对象
                Sheet sheet = wb.getSheet(index);
                // sheet.getRows()返回该页的总行数
                for (int i = 0; i < sheet.getRows(); i++) {
                    List innerList=new ArrayList();
                    // sheet.getColumns()返回该页的总列数
                    for (int j = 0; j < sheet.getColumns(); j++) {
                        String cellinfo = sheet.getCell(j, i).getContents();
                        if(cellinfo.isEmpty()){
                            continue;
                        }
                        innerList.add(cellinfo);
                        System.out.print(cellinfo);
                    }
                    outerList.add(i, innerList);
                    System.out.println();
                }
                return outerList;
            }
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (BiffException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
        return null;
    }
}
