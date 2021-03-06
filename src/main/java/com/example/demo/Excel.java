package com.example.demo;


import org.apache.commons.io.FileUtils;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.Hyperlink;
import org.apache.poi.ss.usermodel.Sheet;
import org.springframework.util.StringUtils;

import java.io.*;
import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.List;

/**
 * Created with IntelliJ IDEA.
 * Author: fanqiuhang
 * Date: 2018/8/27 9:09
 */
public class Excel {
    private static List<Building> list = new ArrayList<>();

    private static HSSFWorkbook wb = new HSSFWorkbook();

    private static HSSFSheet sheet_data = null;

    private static int i = 0;

    public static void export(List<Building> buildings,File phone){

        try {
            HSSFWorkbook wb_phone = new HSSFWorkbook(new FileInputStream(phone));
            sheet_data = wb_phone.getSheetAt(0);
        } catch (IOException e) {
            e.printStackTrace();
        }

        list = buildings;
        /**
         * 创建目录
         */
         for (Building building : list){
             make(building);
         }
        /**
         * 创建房间
         */
        for (Building building :list){
            String name = building.getBuildingNo()+building.getPartNo();
            Integer floor = building.getFloor();
            Integer num = building.getNum();
            for (int i = 0; i < floor; i++) {
                for (int j = 0; j < num; j++) {
                    /*String fore = "";
                    if (i == 13){
                        fore = "12A";
                    } else if (i == 14){
                        fore = "13B";
                    } else if (i == 4){
                        fore = "3A";
                    } else {
                        fore = i + 1 + "";
                    }*/
                    String fore = i + 1 +"";

                    if (i < 9) {
                        fore = "0" + fore;
                    }

                    String end = j + 1 + "";
                    if (j <= 10){
                        end = "0" + (j + 1);
                    }
                    HSSFSheet sheet = wb.createSheet(name + fore + end);
                    init(sheet);
                }
            }
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
            if (cell_back != null && "返回".equals(cell_back.getStringCellValue())){
                setBackStyle(cell_back);

                Integer last = sheet.getLastRowNum() + 1;
                for (int j = 0; j < last; j++) {
                    HSSFRow row = sheet.getRow(j);
                    for (int k = 0; k < 6; k++) {
                        HSSFCell cell = row.getCell(k);
                        if (cell != null) {
                            setCommonStyle(cell);
                        }
                    }
                }
            }
        }
        /**
         * 设置超链接
         */
        for (Building building :list){
            String name = building.getBuildingNo()+building.getPartNo();
            Integer floor = building.getFloor();
            Integer num = building.getNum();
            HSSFSheet sheet = wb.getSheet(name);
            for (int i = 0; i < floor; i++) {
                HSSFRow row = sheet.getRow(i);
                for (int j = 0; j < num; j++) {
                    HSSFCell cell = row.getCell(j);

                    /*String fore = "";
                    if (i == 13){
                        fore = "12A";
                    } else if (i == 14){
                        fore = "13B";
                    } else if (i == 4){
                        fore = "3A";
                    } else {
                        fore = i + 1 + "";
                    }*/
                    String fore = i + 1 + "";
                    if (i < 9) {
                        fore = "0" + fore;
                    }

                    String end = j + 1 + "";
                    if (j <= 10){
                        end = "0" + (j + 1);
                    }

                    Hyperlink hyperlink = new HSSFHyperlink(Hyperlink.LINK_DOCUMENT);
                    String des = name + fore + end;
                    hyperlink.setAddress("#'" + des +"'!A1");
                    cell.setHyperlink(hyperlink);
                }
            }
        }


        try {
            File file = new File("F://test.xls");
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

    /******************************************************************************************************/
    /******************************************************************************************************/
    /******************************************************************************************************/

    /**
     * 创建目录
     * @param building
     */
    private static void make(Building building) {
        HSSFSheet dir = wb.createSheet(building.getBuildingNo()+building.getPartNo());
        Integer floor = building.getFloor();
        Integer num = building.getNum();

        for (int i = 0; i < floor; i++) {
            HSSFRow row = dir.createRow(i);
            for (int j = 0; j < num; j++) {
                HSSFCell cell = row.createCell(j);
                /*String fore = "";
                if (i == 13){
                    fore = "12A";
                } else if (i == 14){
                    fore = "13B";
                } else if (i == 4){
                    fore = "3A";
                } else {
                    fore = i + 1 + "";
                }*/
                String fore = i + 1 + "";

                String end = j + 1 + "";
                if (j <= 8){
                    end = "0" + (j + 1);
                }
                cell.setCellValue(fore + end);
            }
        }
    }

    /**
     * 设置每一个工作簿的内容。当然每个工作簿的内容是相同的
     * @param sheet
     */
    private static void init(HSSFSheet sheet){
        sheet.setColumnWidth(0,30*256);
        sheet.setColumnWidth(1,30*256);
        sheet.setColumnWidth(2,40*256);
        sheet.setColumnWidth(3,30*256);
        sheet.setColumnWidth(4,30*256);
        sheet.setColumnWidth(5,30*256);

        HSSFRow row = sheet.createRow(0);

        HSSFCell c0 = row.createCell(0);
        c0.setCellValue("业主");
        HSSFCell cell_name = row.createCell(1);

        HSSFCell c2 = row.createCell(2);
        c2.setCellValue("座机");
        HSSFCell cell_tel = row.createCell(3);

        HSSFCell c4 = row.createCell(4);
        c4.setCellValue("跟进日期");
        HSSFCell c5 = row.createCell(5);
        c5.setCellValue("沟通内容");

        HSSFCell c6 = row.createCell(6);
        c6.setCellValue("返回");
        Hyperlink hyperlink = new HSSFHyperlink(Hyperlink.LINK_DOCUMENT);
        String sheetName = sheet.getSheetName();
        Integer position = sheetName.indexOf("#");
        String des = sheetName.substring(0,position + 2);
        hyperlink.setAddress("#'" + des +"'!A1");
        c6.setHyperlink(hyperlink);


        HSSFRow row2 = sheet.createRow(1);
        set(row2,"联系方式","其他联系方式");
        HSSFCell cell_phone = row2.getCell(1);


        HSSFRow row3 = sheet.createRow(2);
        set(row3,"代理人","业主爱人");

        HSSFRow row4 = sheet.createRow(3);
        set(row4,"联系方式","业主爱人电话");

        HSSFRow row5 = sheet.createRow(4);
        set(row5,"业主基本情况","房子基本情况");

        HSSFRow row6 = sheet.createRow(5);
        set(row6,"面积","户型特点");
        HSSFCell cell_square = row6.getCell(1);


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

        getData(cell_name,cell_tel,cell_phone,cell_square);
    }

    /**
     * 设置一列的内容，s1放到第一列，s2放到第三列
     * @param row
     * @param s1
     * @param s2
     */
    private static void set(HSSFRow row, String s1,String s2){
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

    /**
     * 设置返回字体样式
     */
    private static HSSFCellStyle cellStyle_back = wb.createCellStyle();
    private static HSSFFont font_back = wb.createFont();
    private static void setBackStyle(HSSFCell cell){
        font_back.setColor(HSSFColor.BLUE.index);
        font_back.setFontHeightInPoints((short) 20);
        font_back.setUnderline((byte) 1);
        cellStyle_back.setFont(font_back);
        cell.setCellStyle(cellStyle_back);
    }

    /**
     * 设置通用字体样式
     */
    private static HSSFCellStyle cellStyle_common = wb.createCellStyle();
    private static HSSFFont font_common = wb.createFont();
    private static void setCommonStyle(HSSFCell cell){
        cellStyle_common.setWrapText(true);
        cellStyle_common.setAlignment((short) 0);
        font_common.setFontHeightInPoints((short) 16);
        cellStyle_common.setFont(font_common);
        cellStyle_common.setDataFormat((short) 0x31);
        cell.setCellStyle(cellStyle_common);
    }


    /**
     * 设置名字，座机，手机，面积 信息
     */
    private static void getData(HSSFCell cell_name,HSSFCell cell_tel,HSSFCell cell_phone,HSSFCell cell_square) {
        HSSFRow row_data = sheet_data.getRow(i);
        System.out.println("第" + (i + 1) + "个信息");
        if (row_data != null) {
            HSSFSheet sheet = cell_name.getSheet();
            String sheet_name = sheet.getSheetName();
            System.out.println(sheet_name);
            StringBuffer buffer = new StringBuffer();
            buffer.append(sheet_name.substring(0,1));
            buffer.append(sheet_name.substring(2,3));
            buffer.append("-");
            buffer.append(sheet_name.substring(3));
            String str = buffer.toString();

            HSSFCell cell_1 = row_data.getCell(1);
            System.out.println(cell_1.getStringCellValue());

            if (str.equals(cell_1.getStringCellValue())) {
                System.out.println("相等");
                //设置名字
                getType(row_data.getCell(2),cell_name);
                //设置座机
                getType(row_data.getCell(4),cell_tel);
                //设置电话
                getType(row_data.getCell(5),cell_phone);
                //设置面积
                String square = String.valueOf(new BigDecimal(row_data.getCell(3).getNumericCellValue()));
                if (square.length() > 7) {
                    square = square.substring(0,6);
                }
                cell_square.setCellValue(square);

                //只有成功读取了该条信息，才能到下一条
                i++;
            }
        }
    }

    private static void getType(HSSFCell cell_data,HSSFCell cell) {
        int type = cell_data.getCellType();
        switch (type) {
            case 0:
                String str = new BigDecimal(cell_data.getNumericCellValue()).toString();
                if (str.indexOf(".") > 0) {
                    str = str.substring(0,str.indexOf("."));
                }
                cell.setCellValue(str);
                break;
            case 1:
                String str2 = cell_data.getStringCellValue();
                cell.setCellValue(str2);
                break;
            default:
                cell.setCellValue("");
        }

    }
}
