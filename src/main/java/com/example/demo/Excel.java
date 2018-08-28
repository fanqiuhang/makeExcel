package com.example.demo;


import org.apache.commons.io.FileUtils;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.Hyperlink;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Arrays;
import java.util.List;

/**
 * Created with IntelliJ IDEA.
 * Author: fanqiuhang
 * Date: 2018/8/27 9:09
 */
public class Excel {
    public static void main(String[] args) {
        export();
    }


    public static void export(){
        HSSFWorkbook wb = new HSSFWorkbook();

        /**
         * 设置返回字体样式
         */
        HSSFCellStyle cellStyle_back = wb.createCellStyle();
        HSSFFont font_back = wb.createFont();
        font_back.setColor(HSSFColor.BLUE.index);
        font_back.setFontHeightInPoints((short) 20);
        font_back.setUnderline((byte) 1);
        cellStyle_back.setFont(font_back);

        /**
         * 设置通用字体样式
         */
        HSSFCellStyle cellStyle_common = wb.createCellStyle();
        cellStyle_common.setWrapText(true);
        cellStyle_common.setAlignment((short) 0);
        HSSFFont font_common = wb.createFont();
        font_common.setFontHeightInPoints((short) 16);
        cellStyle_common.setFont(font_common);
        cellStyle_common.setDataFormat((short) 0x31);

        /**
         * 所有门牌号
         */
        String[] strings = {"1#","5#","1#B101","1#B102","1#B103","1#B104","1#B105","1#B106","1#B107"
                           ,"1#B201","1#B202","1#B203","1#B204","1#B205","1#B206","1#B207"
                           ,"1#B301","1#B302","1#B303","1#B304","1#B305","1#B306","1#B307"
                           ,"1#B3A01","1#B3A02","1#B3A03","1#B3A04","1#B3A05","1#B3A06","1#B3A07"
                           ,"1#B501","1#B502","1#B503","1#B504","1#B505","1#B506","1#B507"
                           ,"1#B601","1#B602","1#B603","1#B604","1#B605","1#B606","1#B607"
                           ,"1#B701","1#B702","1#B703","1#B704","1#B705","1#B706","1#B707"
                           ,"1#B801","1#B802","1#B803","1#B804","1#B805","1#B806","1#B807"
                           ,"1#B901","1#B902","1#B903","1#B904","1#B905","1#B906","1#B907"
                           ,"1#B1001","1#B1002","1#B1003","1#B1004","1#B1005","1#B1006","1#B1007"
                           ,"1#B1101","1#B1102","1#B1103","1#B1104","1#B1105","1#B1106","1#B1107"
                           ,"1#B1201","1#B1202","1#B1203","1#B1204","1#B1205","1#B1206","1#B1207"
                           ,"1#B12A01","1#B12A02","1#B12A03","1#B12A04","1#B12A05","1#B12A06","1#B12A07"
                           ,"1#B12B01","1#B12B02","1#B12B03","1#B12B04","1#B12B05","1#B12B06","1#B12B07"

                ,"5#A101","5#A102","5#A103","5#A104","5#A105","5#A106","5#A107","5#A108"
                ,"5#A201","5#A202","5#A203","5#A204","5#A205","5#A206","5#A207","5#A208"
                ,"5#A301","5#A302","5#A303","5#A304","5#A306","5#A307","5#A308"
                ,"5#A3A01","5#A3A02","5#A3A03","5#A3A04","5#A3A05","5#A3A06","5#A3A07","5#A3A08"
                ,"5#A501","5#A502","5#A503","5#A504","5#A506","5#A507","5#A508"
                ,"5#A601","5#A602","5#A603","5#A604","5#A605","5#A606","5#A607","5#A608"
                ,"5#A701","5#A702","5#A703","5#A704","5#A706","5#A707","5#A708"
                ,"5#A801","5#A802","5#A803","5#A804","5#A805","5#A806","5#A807","5#A808"
                ,"5#A901","5#A902","5#A903","5#A904","5#A906","5#A907","5#A908"
                ,"5#A1001","5#A1002","5#A1003","5#A1004","5#A1005","5#A1006","5#A1007","5#A1008"
                ,"5#A1101","5#A1102","5#A1103","5#A1104","5#A1106","5#A1107","5#A1108"
                ,"5#A1201","5#A1202","5#A1203","5#A1204","5#A1205","5#A1206","5#A1207","5#A1208"
                ,"5#A12A01","5#A12A02","5#A12A03","5#A12A04","5#A12A06","5#A12A07","5#A12A08"
                ,"5#A12B01","5#A12B02","5#A12B03","5#A12B04","5#A12B05","5#A12B06","5#A12B07","5#A12B08"
                ,"5#A1501","5#A1502","5#A1503","5#A1504","5#A1506","5#A1507","5#A1508"
                ,"5#A1601","5#A1602","5#A1603","5#A1604","5#A1605","5#A1606","5#A1607","5#A1608"
                ,"5#A1701","5#A1702","5#A1703","5#A1704","5#A1706","5#A1707","5#A1708"
                ,"5#A1801","5#A1802","5#A1803","5#A1804","5#A1805","5#A1806","5#A1807","5#A1808"
                ,"5#A1901","5#A1902","5#A1903","5#A1904","5#A1906","5#A1907","5#A1908"
                ,"5#A2001","5#A2002","5#A2003","5#A2004","5#A2005","5#A2006","5#A2007","5#A2008"
        };

        List<String> list = Arrays.asList(strings);
        for (String s :list){
            HSSFSheet sheet = wb.createSheet(s);
            sheet.setColumnWidth(0,30*256);
            sheet.setColumnWidth(1,30*256);
            sheet.setColumnWidth(2,40*256);
            sheet.setColumnWidth(3,30*256);
            sheet.setColumnWidth(4,30*256);
            sheet.setColumnWidth(5,30*256);
            init(sheet);
        }

        /**
         * 设置全部单元格格式
         */
        int length = wb.getNumberOfSheets();
        for (int i = 0; i < length; i++) {
            HSSFSheet sheet = wb.getSheetAt(i);
            //设置返回字体样式
            HSSFRow row_back = sheet.getRow(0);
            HSSFCell cell_back = row_back.getCell(6);
            cell_back.setCellStyle(cellStyle_back);

            int last = sheet.getLastRowNum() + 1;
            for (int j = 0; j < last; j++) {
                HSSFRow row = sheet.getRow(j);
                for (int k = 0; k < 6; k++) {
                    HSSFCell cell = row.getCell(k);
                    if (cell != null) {
                        cell.setCellStyle(cellStyle_common);
                    }
                }
            }
        }


        try {
            File file = new File("F://资料.xls");
            file.createNewFile();
            FileOutputStream outputStream = FileUtils.openOutputStream(file);
            wb.write(outputStream);
            outputStream.flush();
            outputStream.close();
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    public static void init(HSSFSheet sheet){
        HSSFRow row = sheet.createRow(0);

        HSSFCell c0 = row.createCell(0);
        c0.setCellValue("业主");
        HSSFCell c1 = row.createCell(1);
        HSSFCell c2 = row.createCell(2);
        c2.setCellValue("座机");
        HSSFCell c3 = row.createCell(3);
        HSSFCell c4 = row.createCell(4);
        c4.setCellValue("跟进日期");
        HSSFCell c5 = row.createCell(5);
        c5.setCellValue("沟通内容");

        HSSFCell c6 = row.createCell(6);
        c6.setCellValue("返回");
        Hyperlink hyperlink = new HSSFHyperlink(Hyperlink.LINK_DOCUMENT);
        String sheetName = sheet.getSheetName();
        if (sheetName.startsWith("1#B")){
            hyperlink.setAddress("#'1#'!G4");
        } else if (sheetName.startsWith("5#A")) {
            hyperlink.setAddress("#'5#'!A4");
        }
        c6.setHyperlink(hyperlink);


        HSSFRow row2 = sheet.createRow(1);
        set(row2,"联系方式","其他联系方式");

        HSSFRow row3 = sheet.createRow(2);
        set(row3,"代理人","业主爱人");

        HSSFRow row4 = sheet.createRow(3);
        set(row4,"联系方式","业主爱人电话");

        HSSFRow row5 = sheet.createRow(4);
        set(row5,"业主基本情况","房子基本情况");

        HSSFRow row6 = sheet.createRow(5);
        set(row6,"面积","户型特点");

        HSSFRow row7 = sheet.createRow(6);
        set(row7,"朝向",null);

        HSSFRow row8 = sheet.createRow(7);
        set(row8,"装修","有无车位");

        HSSFRow row9 = sheet.createRow(8);
        set(row9,"户型","车位号");

        HSSFRow row10 = sheet.createRow(9);
        set(row10,"房屋状态","置换意向小区");

        HSSFRow row11 = sheet.createRow(10);
        set(row11,"租户",null);

        HSSFRow row12 = sheet.createRow(11);
        set(row12,"租户联系方式","是否看过房、什么小区");

        HSSFRow row13 = sheet.createRow(12);
        set(row13,"上一业主姓名",null);

        HSSFRow row14 = sheet.createRow(13);
        set(row14,"上一业主电话","是否转介绍客户");

        HSSFRow row15 = sheet.createRow(14);
        set(row15,"出售记录",null);

    }

    /**
     * 设置一列的内容，s1放到第一列，s2放到第三列
     * @param row
     * @param s1
     * @param s2
     */
    static void set(HSSFRow row, String s1,String s2){
        HSSFCell c0 = row.createCell(0);
        c0.setCellValue(s1);
        HSSFCell c1 = row.createCell(1);
        HSSFCell c2 = row.createCell(2);
        if (s2 != null) {
            c2.setCellValue(s2);
        }
        HSSFCell c3 = row.createCell(3);
        HSSFCell c4 = row.createCell(4);
        HSSFCell c5 = row.createCell(5);

    }
}
