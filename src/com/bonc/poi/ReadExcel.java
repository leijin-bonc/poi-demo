/*
 * 文件名：ReadExcel.java
 * 版权：Copyright by www.bonc.com.cn
 * 描述：
 * 修改人：Jingege
 * 修改时间：2017年7月24日
 */

package com.bonc.poi;

import java.io.File;
import java.io.FileInputStream;
import java.text.SimpleDateFormat;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;


/**
 * 读取excel内容,包括xls，xlsx格式
 * @author Jingege
 * @version 2017年7月24日
 * @see ReadExcel
 * @since
 */
public class ReadExcel{
    /**
     * Description: <br>
     *  读取excel内容
     * @return CellValue
     * @see
     */
    public String readExcel(){
        SimpleDateFormat fmt = new SimpleDateFormat("yyyy-mm-dd");
        try{
           // File excelFile1 = new File("D:\\BONC\\project\\test.xls"); //创建文件对象
            File excelFile = new File("D:\\BONC\\project\\xlsx.xlsx");//创建xmls
            FileInputStream is = new FileInputStream(excelFile);//文件流
            Workbook workbook = WorkbookFactory.create(is);  //这种方式处理2003,2007,2010，13即.xls或.xlsx的文件
            
            int sheetCount = workbook.getNumberOfSheets();//sheet的数量
            for (int s = 0; s < sheetCount; s ++) { //遍历sheet
                Sheet sheet = workbook.getSheetAt(s);
                int rowCount = sheet.getPhysicalNumberOfRows();
                for(int r = 0; r < rowCount; r ++) {
                    Row row = sheet.getRow(r);
                    int cellCount = row.getPhysicalNumberOfCells();
                    for (int c = 0; c < cellCount; c ++) {
                        Cell cell = row.getCell(c);
                        int cellType = cell.getCellType();
                        String cellValue = null;
                        switch(cellType) {
                            case Cell.CELL_TYPE_STRING://文本
                                cellValue = cell.getStringCellValue();
                                break;
                            case Cell.CELL_TYPE_NUMERIC://数字，日期
                                if (DateUtil.isCellDateFormatted(cell)){
                                    cellValue = fmt.format(cell.getDateCellValue());
                                } else {
                                    cellValue = String.valueOf(cell.getNumericCellValue());
                                }
                                break;
                            case Cell.CELL_TYPE_BOOLEAN://布尔型
                                cellValue = String.valueOf(cell.getBooleanCellValue());
                                break;
                            case Cell.CELL_TYPE_BLANK://空白
                                cellValue = cell.getStringCellValue();
                                break;
                            case Cell.CELL_TYPE_ERROR://错误
                                cellValue = "错误";
                                break;
                            case Cell.CELL_TYPE_FORMULA://公式
                                cellValue = "错误";
                                break;
                            default:
                                cellValue = "错误";
                        }
                        System.out.println(cellValue + "    ");
                    }
                    System.out.println();
                }
            }
        } catch (Exception e){
            e.printStackTrace();
        }
        return "";
    }

}
