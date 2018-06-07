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
public class ReadExcel {
    public static void main(String[] args) {
        ReadExcel obj = new ReadExcel();
        // 此处为我创建Excel路径：E:/zhanhj/studysrc/jxl下
        //File file = new File("D:\\0_1\\蚂蚁财富号\\20180426_每日财经和问答红包\\问答红包\\题库2018.4.9-筛选.xls");
        File file = new File("D:\\0_1\\蚂蚁财富号\\20180525_简单红包问答\\题目更新.xls");
        List excelList = obj.readExcel(file);
        String sql = "INSERT INTO question_list_info (module,question_no,question,answer,answer_parse,options) VALUES ('04',(SELECT a.num from (SELECT (MAX(question_no)+1)num FROM question_list_info) a ),";
        String deletSql = "DELETE FROM question_list_info WHERE question = ";
        String deletSqlEnd = ";";
        String sqlend = ");";
        String nextLineStr = "\r\n";
        StringBuffer insertSql = new StringBuffer();
        StringBuffer delSql = new StringBuffer();
        System.out.println("list中的数据打印出来");
        for (int i = 1; i < excelList.size(); i++) {
        	
            List list = (List) excelList.get(i);
            StringBuffer sqlValue = new StringBuffer();
        	StringBuffer opptions = new StringBuffer();
        	
        	StringBuffer DelsqlValue = new StringBuffer();
            for (int j = 0; j < list.size(); j++) {
            	
            	/*if(j == 3){
            		opptions.append("'"+ "A:A:"+list.get(j));
            		//System.out.print("A:"+list.get(j));
            	}else if(j == 4){
            		opptions.append( ";B:B:"+list.get(j));
            		//System.out.print("B:"+list.get(j));
            	}else if(j == 5){
            		if(!"  ".equals(list.get(j).toString())){
            			opptions.append(";C:C:"+list.get(j)+"'");
            		}else{
            			opptions.append("'");
            		}
            		
            		//System.out.print("C:"+list.get(j));
            	}else if(j == 6){
            		if("1".equals(list.get(j).toString())){
            			sqlValue.append("'A',");
            			//System.out.print("A");
            		}else if("2".equals(list.get(j).toString())){
            			sqlValue.append("'B',");
            			//System.out.print("B");
            		}else if("3".equals(list.get(j).toString())){
            			sqlValue.append("'C',");
            			//System.out.print("C");
            		}            		
            	}else if(j == 2){
            		
            	}else {
            		sqlValue.append("'"+list.get(j)+"',");
            		//System.out.print(list.get(j));
            	}*/
            	//简单红包
            	if(j == 0){
            		DelsqlValue.append("'"+list.get(j)+"'");
            	}
            	if(j == 1){
            		opptions.append("'"+ "A:"+list.get(j));
            		//System.out.print("A:"+list.get(j));
            	}else if(j == 2){
            		if(!"  ".equals(list.get(j).toString())){
            			opptions.append(";B:"+list.get(j)+"'");
            		}else{
            			opptions.append("'");
            		}
            		
            		//System.out.print("C:"+list.get(j));
            	}else if(j == 3){
            		if("1".equals(list.get(j).toString())){
            			sqlValue.append("'A',");
            			//System.out.print("A");
            		}else if("2".equals(list.get(j).toString())){
            			sqlValue.append("'B',");
            			//System.out.print("B");
            		}else if("3".equals(list.get(j).toString())){
            			sqlValue.append("'C',");
            			//System.out.print("C");
            		}else{
            			sqlValue.append("'").append(list.get(j).toString()).append("',");
            		}
            		
            	}else {
            		sqlValue.append("'"+list.get(j)+"',");
            		//System.out.print(list.get(j));
            	}
            	
                
            }
            System.err.println(deletSql.concat(DelsqlValue.toString()).concat(deletSqlEnd));
            System.err.println(sql.concat(sqlValue.toString()).concat(opptions.toString()).concat(sqlend));
            insertSql.append(sql.concat(sqlValue.toString()).concat(opptions.toString()).concat(sqlend)).append(nextLineStr);
            delSql.append(deletSql.concat(DelsqlValue.toString()).concat(deletSqlEnd)).append(nextLineStr);
        }
        
        System.err.println("=====================================");
        System.err.println(delSql.toString());
        System.err.println(insertSql.toString());
        try {
			OutputStream fos=new FileOutputStream("D:\\0_1\\蚂蚁财富号\\20180525_简单红包问答\\20180530_dml_question_info.sql");
			fos.write(delSql.toString().getBytes());
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
