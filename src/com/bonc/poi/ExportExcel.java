/*
 * 文件名：ExportExcel.java
 * 版权：Copyright by www.bonc.com.cn
 * 描述：
 * 修改人：Jingege
 * 修改时间：2017年7月24日
 */

package com.bonc.poi;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.util.CellRangeAddress;

public class ExportExcel{
    public void exportExcel()throws IOException{
        HSSFWorkbook wb = new HSSFWorkbook(); //创建工作簿
        //创建工作表
        HSSFSheet sheet = wb.createSheet("new sheet");
        for (int i = 0;i < 3; i ++) {
            //设置列宽
            sheet.setColumnWidth(i, 3000);
        }
        //创建行
        HSSFRow row = sheet.createRow(0);
        row.setHeightInPoints(30);//设置行高
        //创建单元格
        HSSFCell cell = row.createCell(0);
        cell.setCellValue("用户信息表");
        
        //标题样式
        //创建单元格样式
        HSSFCellStyle cellStyle = wb.createCellStyle();
        //设置单元格的背景颜色为淡蓝色
        cellStyle.setFillForegroundColor(HSSFColor.PALE_BLUE.index);
        cellStyle.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
        //设置单元格垂直居中对齐
        cellStyle.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER);
        //设置单元格居中对齐
        cellStyle.setAlignment(HSSFCellStyle.ALIGN_CENTER);
        //设置单元格内容显示不下自动换行
        cellStyle.setWrapText(true);
        //设置单元格字体样式
        HSSFFont font = wb.createFont();
        //设置字体加粗
        font.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);
        font.setFontName("宋体");
        font.setFontHeight((short) 200);
        cellStyle.setFont(font);
        //设置单元格边框为细线条  
        cellStyle.setBorderLeft(HSSFCellStyle.BORDER_THIN);  
        cellStyle.setBorderBottom(HSSFCellStyle.BORDER_THIN);  
        cellStyle.setBorderRight(HSSFCellStyle.BORDER_THIN);  
        cellStyle.setBorderTop(HSSFCellStyle.BORDER_THIN); 
        //设置单元格样式
        cell.setCellStyle(cellStyle);
        //合并单元格
        sheet.addMergedRegion(new CellRangeAddress(0,0,0,2));//合并0行2列
        
        HSSFRow row1 = sheet.createRow(1);
        //标题信息
        String[] titles = {"ID","用户名","密码"};
        for (int i = 0;i < 3; i ++) {
            HSSFCell cell1 = row1.createCell(i);
            cell1.setCellValue(titles[i]);
            //设置单元格样式
            cell1.setCellStyle(cellStyle);
        }
      //模拟数据，实际情况下String[]多为实体bean  
        List<String[]> list = new ArrayList<String[]>();  
        list.add(new String[]{"1","张三","111"});  
        list.add(new String[]{"2","李四","222"});  
        list.add(new String[]{"3","王五","333"});
        
        //内容样式
        //创建单元格样式
        HSSFCellStyle cellStyle2 = wb.createCellStyle();  
        // 设置单元格居中对齐  
        cellStyle2.setAlignment(HSSFCellStyle.ALIGN_CENTER);  
        // 设置单元格垂直居中对齐  
        cellStyle2.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER);  
        // 创建单元格内容显示不下时自动换行  
        cellStyle2.setWrapText(true);  
        // 设置单元格字体样式  
        HSSFFont font2 = wb.createFont();  
        // 设置字体加粗  
        font2.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);  
        font2.setFontName("宋体");  
        font2.setFontHeight((short) 200);  
        cellStyle2.setFont(font2);  
        // 设置单元格边框为细线条  
        cellStyle.setBorderLeft(HSSFCellStyle.BORDER_THIN);  
        cellStyle.setBorderBottom(HSSFCellStyle.BORDER_THIN);  
        cellStyle.setBorderRight(HSSFCellStyle.BORDER_THIN);  
        cellStyle.setBorderTop(HSSFCellStyle.BORDER_THIN);  
        //循环赋值  
        for(int i=0;i<list.size();i++){  
            HSSFRow row2 = sheet.createRow(i+2);  
            for(int j=0;j<3;j++){  
                HSSFCell cell1 = row2.createCell(j);  
                cell1.setCellValue(list.get(i)[j]);  
                //设置单元格样式  
                cell1.setCellStyle(cellStyle2);  
            }  
        }  
        File file = new File("D://a.xls");  
        if(!file.exists()){  
            file.createNewFile();  
        }  
        FileOutputStream fileOut = new FileOutputStream(file);  
        wb.write(fileOut);  
        fileOut.close();
    }

}
