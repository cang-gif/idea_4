package com.ht.T2;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hssf.util.Region;
import org.apache.poi.ss.usermodel.Font;

import java.io.*;

public class ExcelOutput2 {
    public static void main(String[] args) {
        HSSFWorkbook workbook = new HSSFWorkbook();
        HSSFSheet sheet = workbook.createSheet();
        HSSFRow row = sheet.createRow(0);

        HSSFCell cel0 = row.createCell(0);
        cel0.setCellValue("成绩单");

        HSSFRow row0 = sheet.createRow(1);
        HSSFCell cell0 = row0.createCell(0);
        cell0.setCellValue("编号");

        HSSFCell cell1 = row0.createCell(1);
        cell1.setCellValue("姓名");

        HSSFCell cell2 = row0.createCell(2);
        cell2.setCellValue("语文");

        HSSFCell cell3 = row0.createCell(3);
        cell3.setCellValue("数学");

        HSSFCell cell4 = row0.createCell(4);
        cell4.setCellValue("英语");

        HSSFCell cell5 = row0.createCell(5);
        cell5.setCellValue("总分");

        int i=0;
        for(i=2; i < 8;i++){
            HSSFRow rows = sheet.createRow(i);
            HSSFCell cells0 = rows.createCell(0);
            cells0.setCellValue(i-1);

            HSSFCell cells1 = rows.createCell(1);
            cells1.setCellValue("张三"+(i-1));

            HSSFCell cells2 = rows.createCell(2);
            cells2.setCellValue(10*(i-1));

            HSSFCell cells3 = rows.createCell(3);
            cells3.setCellValue(10*(i-1));

            HSSFCell cells4 = rows.createCell(4);
            cells4.setCellValue(10*(i-1));

            HSSFCell cells5 = rows.createCell(5);
            cells5.setCellFormula("sum(c"+(i+1)+":e"+(i+1)+")");
        }

        HSSFRow rows = sheet.createRow(i);
        HSSFCell cells0 = rows.createCell(0);
        cells0.setCellValue("总计");

//        HSSFCell cells1 = rows.createCell(1);
//        cells1.setCellValue("张三"+(i-1));

        HSSFCell cells2 = rows.createCell(2);
        cells2.setCellFormula("sum(c3:c"+i+")");

        HSSFCell cells3 = rows.createCell(3);
        cells3.setCellFormula("sum(d3:d"+i+")");

        HSSFCell cells4 = rows.createCell(4);
        cells4.setCellFormula("sum(e3:e"+i+")");

        HSSFCell cells5 = rows.createCell(5);
        cells5.setCellFormula("sum(f3:f"+i+")");

        Region region0 = new Region(0, (short) 0, 0, (short) 5);
        sheet.addMergedRegion(region0);
        //行高
        row.setHeightInPoints(50);
        //
        HSSFCellStyle style = workbook.createCellStyle();
        style.setAlignment(HSSFCellStyle.ALIGN_CENTER);//水平居中
        style.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER);//垂直居中
        HSSFFont font = workbook.createFont();
        font.setFontHeightInPoints((short)18);//字体大小
        font.setBoldweight(Font.BOLDWEIGHT_BOLD);//加粗
        style.setFont(font);
        cel0.setCellStyle(style);

        try {
            OutputStream stream = new FileOutputStream("E:\\Work.xls");
            workbook.write(stream);
            System.out.println("操作成功");
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }

    }
}
