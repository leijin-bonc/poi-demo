/*
 * 文件名：ExcelTest.java
 * 版权：Copyright by www.bonc.com.cn
 * 描述：
 * 修改人：Jingege
 * 修改时间：2017年7月24日
 */

package com.bonc.poi;

import java.io.IOException;

public class ExcelTest
{

    public static void main(String[] args) throws IOException
    {
        // TODO Auto-generated method stub
        ReadExcel readExcel = new ReadExcel();
        //readExcel.readExcel();
        ExportExcel exportExcel = new ExportExcel();
        exportExcel.exportExcel();
    }

}
