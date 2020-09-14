package com.ht.test;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;

public class ExccelInput1 {
    public static void main(String[] args) throws IOException {
        //得到文件输入流
        InputStream inputStream = new FileInputStream(new File("E:\\学习资料\\Excel（POI）+Quartz+Open CSV\\T1\\test.xls"));
        //通过文件输入流得到Excel文档对象
        HSSFWorkbook hssfWorkbook = new HSSFWorkbook(inputStream);
        //获取工作表个数，循环工作表
        for (int i = 0; i < hssfWorkbook.getNumberOfSheets(); i++) {
            HSSFSheet hssfSheet = hssfWorkbook.getSheetAt(i);
            System.out.println("工作表名称： "+hssfSheet.getSheetName());
            System.out.println("最大行数： "+hssfSheet.getLastRowNum());
            if(hssfSheet.getLastRowNum()!=0){
                //获取工作表下的带数据的所有行
                for (int j = 0; j <= hssfSheet.getLastRowNum(); j++) {
                    HSSFRow row =hssfSheet.getRow(j);
                    //获取行下带数据的单元格
                    for (int k = 0; k < row.getLastCellNum(); k++) {
                        HSSFCell cell=  row.getCell(k);
                        System.out.print(cell+" ");
                    }
                    System.out.println();
                }
            }
        }
    }
}
