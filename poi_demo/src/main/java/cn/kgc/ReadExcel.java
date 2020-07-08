package cn.kgc;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * excel文件组成
 *  workbook:工作簿：整个excel
 *      sheet:工作表
 *          row:行
 *              cell:单元格
 */
public class ReadExcel {
    public static void main(String[] args) throws Exception {
        //创建工作簿
        XSSFWorkbook workbook = new XSSFWorkbook("C:\\Users\\25988\\Desktop\\面试题\\poi-demo.xlsx");
//获取工作表，既可以根据工作表的顺序获取，也可以根据工作表的名称获取
        XSSFSheet sheet = workbook.getSheetAt(0);
//遍历工作表获得行对象
        for (Row row : sheet) {
            //遍历行对象获取单元格对象
            for (Cell cell : row) {
                //获得单元格中的值
                String value = cell.getStringCellValue();
                System.out.println(value);
            }
        }
        workbook.close();
    }
}

