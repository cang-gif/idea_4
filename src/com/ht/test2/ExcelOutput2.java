package com.ht.test2;

import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.Region;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;


public class ExcelOutput2 {
    public static void main(String[] args) throws IOException {
        //创建Excel对象
        HSSFWorkbook workbook  = new HSSFWorkbook();
        //创建工作表
        HSSFSheet sheet = workbook.createSheet();
        //第一行
        HSSFRow row0 = sheet.createRow(0);
        row0.createCell(0).setCellValue("工号");
        row0.createCell(1).setCellValue("姓名");
        row0.createCell(2).setCellValue("部门");

        row0.createCell(3).setCellValue("2010年度休假数据");
        row0.createCell(7).setCellValue("2011年度休假数据");

        //第一行
        HSSFRow row1 = sheet.createRow(1);
        row1.createCell(3).setCellValue("2010法定假日");
        row1.createCell(4).setCellValue("2010弹性假日");
        row1.createCell(5).setCellValue("2010病假假日");
        row1.createCell(6).setCellValue("2010补充假日");

        row1.createCell(7).setCellValue("2011法定假日");
        row1.createCell(8).setCellValue("2011弹性假日");
        row1.createCell(9).setCellValue("2011病假假日");
        row1.createCell(10).setCellValue("2011补充假日");

        //设置合并单元格
        sheet.addMergedRegion(new Region(0, (short)0, 1, (short)0));
        sheet.addMergedRegion(new Region(0, (short)1, 1, (short)1));
        sheet.addMergedRegion(new Region(0, (short)2, 1, (short)2));
        sheet.addMergedRegion(new Region(0, (short)3, 0, (short)6));
        sheet.addMergedRegion(new Region(0, (short)7, 0, (short)10));


        OutputStream outputStream =  new FileOutputStream(new File("D://休假数据.xls"));
        workbook.write(outputStream);
        System.out.println("导出成功！");

    }
}
