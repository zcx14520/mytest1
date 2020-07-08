package cn.kgc;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;

/**
 * 向excel中写数据
 */
public class WriteExcel {
    public static void main(String[] args) throws Exception {
        //1.在内存中创建一个excel
        XSSFWorkbook workbook = new XSSFWorkbook();
        //2.在excel工作簿中创建一个工作表
        XSSFSheet sheet = workbook.createSheet("健康管理");
        //3.在工作表中创建单元格
        XSSFRow row = sheet.createRow(0);//第一行
        //4.将数据写道单元格中
        XSSFCell cell = row.createCell(0);
        cell.setCellValue("编号");
        row.createCell(1).setCellValue("名称");
        row.createCell(2).setCellValue("年龄");

        XSSFRow row2 = sheet.createRow(1);//第二行
        row2.createCell(0).setCellValue("110");
        row2.createCell(1).setCellValue("张三");
        row2.createCell(2).setCellValue("28");
        //5.将内存中的工作簿写出到硬盘中
        FileOutputStream out = new FileOutputStream("C:\\Users\\25988\\Desktop\\面试题\\a.xlsx");
        workbook.write(out);
        //6.关闭流
        out.flush();//刷新缓存
        out.close();
        workbook.close();

    }
}
