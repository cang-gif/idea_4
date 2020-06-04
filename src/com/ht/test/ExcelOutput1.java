package com.ht.test;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;

public class ExcelOutput1 {
    public static void main(String[] args) throws IOException {
        //创建Excel对象
        HSSFWorkbook workbook = new HSSFWorkbook();
        //创建工作表
        HSSFSheet sheet = workbook.createSheet();
        //创建行
        HSSFRow row0 = sheet.createRow(0);
        //创建列
        HSSFCell cell0 = row0.createCell(0);
        cell0.setCellValue("编号");
        HSSFCell cell1 = row0.createCell(1);
        cell1.setCellValue("姓名");

        //创建数据行
        HSSFRow row1 = sheet.createRow(1);
        //创建列
        HSSFCell cell00 = row1.createCell(0);
        cell00.setCellValue("1");
        HSSFCell cell01 = row1.createCell(1);
        cell01.setCellValue("李四");
        OutputStream outputStream =  new FileOutputStream(new File("D://Excel导出.xls"));
        workbook.write(outputStream);
        System.out.println("导出成功！");
    }
}
