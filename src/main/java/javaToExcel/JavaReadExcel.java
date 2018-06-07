package javaToExcel;

import java.io.File;
import java.io.FileInputStream;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.HashMap;
import java.util.LinkedList;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class JavaReadExcel {
	
	public static <T> Map<String, List<? extends T>> readExcel(String path, Class clzz) {
        SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd");
        List<T> list = new LinkedList<T>(); 
        Map<String, List<? extends T>> map = new HashMap<String, List<? extends T>>();
        File file = new File(path);
        FileInputStream fis = null;
        Workbook workBook = null;
        if (file.exists()) {
            try {
                fis = new FileInputStream(file); 
                workBook = WorkbookFactory.create(fis);
                int numberOfSheets = workBook.getNumberOfSheets();
                for (int s = 0; s < numberOfSheets; s++) { // sheet工作表
                    Sheet sheetAt = workBook.getSheetAt(s);
//                  String sheetName = sheetAt.getSheetName(); //获取工作表名称
                    int rowsOfSheet = sheetAt.getPhysicalNumberOfRows(); // 获取当前Sheet的总列数
                    System.out.println("当前表格的总行数:" + rowsOfSheet);
                    for (int r = 0; r < rowsOfSheet; r++) { // 总行
                        Row row = sheetAt.getRow(r);
                        if (row == null) {
                            continue;
                        } else {
                            int rowNum = row.getRowNum();
                            System.out.println("当前行:" + rowNum);
                            int numberOfCells = row.getPhysicalNumberOfCells();
                            for (int c = 0; c < numberOfCells; c++) { // 总列(格)
                                Cell cell = row.getCell(c);
                                if (cell == null) {
                                    continue;
                                } else {
                                    int cellType = cell.getCellType();
                                    switch (cellType) {
                                    case Cell.CELL_TYPE_STRING: // 代表文本
                                        String stringCellValue = cell.getStringCellValue();
                                        System.out.print(stringCellValue + "\t");
                                        break;
                                    case Cell.CELL_TYPE_BLANK: // 空白格
                                        String stringCellBlankValue = cell.getStringCellValue();
                                        System.out.print(stringCellBlankValue + "\t");
                                        break;
                                    case Cell.CELL_TYPE_BOOLEAN: // 布尔型
                                        boolean booleanCellValue = cell.getBooleanCellValue();
                                        System.out.print(booleanCellValue + "\t");
                                        break;
                                    case Cell.CELL_TYPE_NUMERIC: // 数字||日期
                                        boolean cellDateFormatted = DateUtil.isCellDateFormatted(cell);
                                        if (cellDateFormatted) {
                                            Date dateCellValue = cell.getDateCellValue();
                                            System.out.print(sdf.format(dateCellValue) + "\t");
                                        } else {
                                            double numericCellValue = cell.getNumericCellValue();
                                            System.out.print(numericCellValue + "\t");
                                        }
                                        break;
                                    case Cell.CELL_TYPE_ERROR: // 错误
                                        byte errorCellValue = cell.getErrorCellValue();
                                        System.out.print(errorCellValue + "\t");
                                        break;
                                    case Cell.CELL_TYPE_FORMULA: // 公式
                                        int cachedFormulaResultType = cell.getCachedFormulaResultType();
                                        System.out.print(cachedFormulaResultType + "\t");
                                        break;
                                    }
                                }
                            }
                            System.out.println(" \t ");
                        }
                        System.out.println("");
                    }
                }
                if (fis != null) {
                    fis.close();
                }
            } catch (Exception e) {
                // TODO Auto-generated catch block
                e.printStackTrace();
            }
        } else {
            System.out.println("文件不存在!");
        }
        return map;
    }
	
	
	public static void main(String[] args) {
		String path = "D:\\0_1\\蚂蚁财富号\\问答红包\\题库2018.4.9-筛选.xlsx";
		JavaReadExcel.readExcel(path, null);		
	}
	
}
