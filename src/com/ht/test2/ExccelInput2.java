package com.ht.test2;

import org.apache.poi.hssf.usermodel.*;

import java.io.*;
import java.text.SimpleDateFormat;
import java.util.Date;

public class ExccelInput2 {
    public static void main(String[] args) throws IOException {
        //得到文件输入流
        InputStream inputStream = new FileInputStream(new File("E:\\学习资料\\Excel（POI）+Quartz+Open CSV\\T1\\成绩表.xls"));
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
                        if(j>0&&k==0){
                        }else if(cell.getCellType()==HSSFCell.CELL_TYPE_FORMULA){//如果是公式
                            System.out.print("公式："+cell+" 值："+cell.getNumericCellValue()+" ");
                        }else if(cell.getCellType()==HSSFCell.CELL_TYPE_NUMERIC&& HSSFDateUtil.isCellDateFormatted(cell)){
                            Date date = cell.getDateCellValue();//获取date类型
                            SimpleDateFormat sDateFormat = new SimpleDateFormat("yyyy-MM-dd");
                            String sdate = sDateFormat.format(date);
                            System.out.print(sdate);
                        }else{//常规单元格类型
                            System.out.print(cell+" ");
                        }
                    }
                    System.out.println();
                }
            }

        }

    }
}
