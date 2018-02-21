package com.funny.office.utils;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.Color;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFColor;

public class ColorUtil{ 

/**
     * excel97中颜色转化为uof颜色
     * 
     * @param color
     *            颜色序号
     * @return 颜色或者null
     */
    public static ColorInfo excel97Color2UOF(Workbook book, short color) {
        if (book instanceof HSSFWorkbook) {
            HSSFWorkbook hb = (HSSFWorkbook) book;
            HSSFColor hc = hb.getCustomPalette().getColor(color);
            ColorInfo ci = excelColor2UOF(hc);
            return ci;
        }
        return null;
    }

    /**
     * excel(包含97和2007)中颜色转化为uof颜色
     * 
     * @param color
     *            颜色序号
     * @return 颜色或者null
     */
    public static ColorInfo excelColor2UOF(Color color) {
        if (color == null) {
            return null;
        }
        ColorInfo ci = null;
        if (color instanceof XSSFColor) {// .xlsx
            XSSFColor xc = (XSSFColor) color;
            byte[] b = xc.getRgb();
            if (b != null) {// 一定是argb
                if(b.length==4)
                ci = ColorInfo.fromARGB(b[0], b[1], b[2], b[3]);
                else
                    ci = ColorInfo.fromARGB(b[0], b[1], b[2]);
            }
        } else if (color instanceof HSSFColor) {// .xls
            HSSFColor hc = (HSSFColor) color;
            short[] s = hc.getTriplet();// 一定是rgb
            if (s != null) {
                ci = ColorInfo.fromARGB(s[0], s[1], s[2]);
            }
        }
        return ci;
    }

    //将颜色转换为16进制的数
    public static String convertColorToHex(Color c){
        ColorInfo ci  = ColorUtil.excelColor2UOF(c);
        String colorStr = null;
        if(ci != null){
            colorStr = OperaColor.toHex(ci.R, ci.G, ci.B);
        }
        return colorStr;
    }
}