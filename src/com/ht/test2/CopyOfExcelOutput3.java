package com.ht.test2;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.Region;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;

public class CopyOfExcelOutput3 {
    public static void main(String[] args) throws IOException {
        //创建Excel对象
        HSSFWorkbook workbook = new HSSFWorkbook();
        //创建工作表
//		workbook.createSheet("员工列表");
        HSSFSheet sheet = workbook.createSheet();

        //标题列
        //创建行
        HSSFRow row = sheet.createRow(0);
        row.createCell(0).setCellValue("成绩表");

        //创建合并单元格对象
        Region region1 = new Region(0, (short) 0, 0, (short) 5);
        sheet.addMergedRegion(region1);

        //创建行
        HSSFRow row0 = sheet.createRow(1);
        //创建列
        HSSFCell cell0 = row0.createCell(0);
        cell0.setCellValue("编号");

        HSSFCell cell1 = row0.createCell(1);
        cell1.setCellValue("姓名");

        HSSFCell cell2 = row0.createCell(2);
        cell2.setCellValue("语文");

        HSSFCell cell3 = row0.createCell(3);
        cell3.setCellValue("数学");

        HSSFCell cell4 = row0.createCell(4);
        cell4.setCellValue("总分");

        HSSFCell cell5 = row0.createCell(5);
        cell5.setCellValue("平均分");

        //创建数据行
        HSSFRow row1 = sheet.createRow(2);
        //创建列
        HSSFCell cell00 = row1.createCell(0);
        cell00.setCellValue("1");

        HSSFCell cell01 = row1.createCell(1);
        cell01.setCellValue("李四");

        HSSFCell cell02 = row1.createCell(2);
        cell02.setCellValue(78);

        HSSFCell cell03 = row1.createCell(3);
        cell03.setCellValue(78);

        //总分(直接写公式 不需要加上等于号)
        //=sum(d2:f2) 范围选择
        //=sum(d2,f2,g4) 点选
        //=(d2+f2) 点选
        HSSFCell cell04 = row1.createCell(4);
        cell04.setCellFormula("(c3+d3)");
        System.out.println();

        //平均分
        HSSFCell cell05 = row1.createCell(5);
        cell05.setCellFormula("(E3/2)");

        OutputStream outputStream = new FileOutputStream(new File("D://Excel导出.xls"));
        workbook.write(outputStream);
        System.out.println("导出成功！");

        }
    }
