package com.funny.office.utils;


import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFFormulaEvaluator;
import org.apache.poi.hssf.usermodel.HSSFPalette;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellValue;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFFormulaEvaluator;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * 	此工具类将EXCEL(XLSX)转换为HTML
 * 依赖包:poi-3.2.jar、poi-ooxml-schemas-3.12.jar、poi-ooxml-3.12.jar 包
 * 需要引入:ColorInfo.java、ColorUtil.java、OperaColor.java
 * @author funnyZpC
 *
 */

public  class Excel2HtmlUtils {
    static String[] bordesr = {"border-top:", "border-right:",
            "border-bottom:", "border-left:"};
    /**
     * 因获取边框只有是或否，故边框类型实际上只取第一个或第二个元素
     */

    static String[] borderStyles = {"solid ", "solid ", "hidden ", "dotted ",
            "none ", "double ", "groove ", "ridge ", "inset ", "outset", "inherit"};
    /*
    static String[] borderStyles= {  "dashed ","solid ", "hidden ", "dotted ",
        "none ", "double ", "groove ", "ridge ", "inset ", "outset", "inherit"};
        */

    /**
     * 转换xls中的颜色代码
     *
     * @param hc
     * @return
     */
    private static String convertToStardColor(HSSFColor hc) {
        StringBuffer sb = new StringBuffer("");
        if (hc != null) {
            if (HSSFColor.AUTOMATIC.index == hc.getIndex()) {
                return null;
            }
            sb.append("#");
            for (int i = 0; i < hc.getTriplet().length; i++) {
                sb.append(fillWithZero(Integer.toHexString(hc.getTriplet()[i])));
            }
        }
        return sb.toString();
    }

    private static String fillWithZero(String str) {
        if (str != null && str.length() < 2) {
            return "0" + str;
        }
        return str;
    }

    /**
     * 获取xls里面的边框
     *
     * @param palette
     * @param b
     * @param s
     * @param t
     * @return
     */
    private static String getBorderStyle(HSSFPalette palette, int b, short s, short t) {
        if (s == 0)
            return String.format("%s %s %s", bordesr[b], borderStyles[s], "#d0d7e5 1px;");
        String borderColorStr = convertToStardColor(palette.getColor(t));
        borderColorStr = borderColorStr == null || borderColorStr.length() < 1 ? "#000000" : borderColorStr;
        return bordesr[b] + borderStyles[s] + borderColorStr + " 1px;";

    }

    /**
     * 获取xlsx里面的边框
     *
     * @param b 0:上 1:右 2:下 3:左
     * @param s 0:未设置当前边框线 1:已设置边框线
     * @param t
     * @return
     */
    private static String getBorderStyle(int b, short s, XSSFColor t) {
    	/*
 			getBorderStyle(0, xcellStyle.getBorderTop(), xcellStyle.getTopBorderXSSFColor());
                */
        if (s == 0) {
            //return bordesr[b] + borderStyles[s] + "#d0d7e5 1px;";
            return String.format("%s%s %s %s", bordesr[b], "1px ", borderStyles[s], "#9facc5;");//默认类似Excel灰线条
        }
        String borderColorStr = ColorUtil.convertColorToHex(t);
        borderColorStr = borderColorStr == null || borderColorStr.length() < 1 ? "#000000" : borderColorStr;
        return String.format("%s%s %s %s; ", bordesr[b], "1px ", borderStyles[s], borderColorStr);
    }

    /**
     * 转换单元格中上中下对齐
     *
     * @param verticalAlignment
     * @return
     */
    private static String convertVerticalAlignToHtml(short verticalAlignment) {
        String valign = "middle";
        switch (verticalAlignment) {
            case XSSFCellStyle.VERTICAL_BOTTOM:
                valign = "bottom";
                break;
            case XSSFCellStyle.VERTICAL_CENTER:
                valign = "center";
                break;
            case XSSFCellStyle.VERTICAL_TOP:
                valign = "top";
                break;
            default:
                break;
        }
        return valign;
    }

    /**
     * 转换单元格中左中右对齐
     *
     * @param alignment
     * @return
     */
    private static String convertAlignToHtml(short alignment) {
        String align = "left";
        switch (alignment) {
            case XSSFCellStyle.ALIGN_LEFT:
                align = "left";
                break;
            case XSSFCellStyle.ALIGN_CENTER:
                align = "center";
                break;
            case XSSFCellStyle.ALIGN_RIGHT:
                align = "right";
                break;
            default:
                break;
        }
        return align;
    }

    /**
     * 空值样式
     *
     * @return
     */
    private static String getNullCellBorderStyle() {
        return "border: #d0d7e5 1px 1px 1px 1px;";
    }

    private static Map<String, String>[] getRowSpanColSpanMap(Sheet sheet) {
        Map<String, String> map0 = new HashMap<String, String>();
        Map<String, String> map1 = new HashMap<String, String>();
        int mergedNum = sheet.getNumMergedRegions();
        CellRangeAddress range = null;
        for (int i = 0; i < mergedNum; i++) {
            range = sheet.getMergedRegion(i);
            int topRow = range.getFirstRow();
            int topCol = range.getFirstColumn();
            int bottomRow = range.getLastRow();
            int bottomCol = range.getLastColumn();
            map0.put(topRow + "," + topCol, bottomRow + "," + bottomCol);
            int tempRow = topRow;
            while (tempRow <= bottomRow) {
                int tempCol = topCol;
                while (tempCol <= bottomCol) {
                    map1.put(tempRow + "," + tempCol, "");
                    tempCol++;
                }
                tempRow++;
            }
            map1.remove(topRow + "," + topCol);
        }
        @SuppressWarnings("rawtypes")
        Map[] map = {map0, map1};
        return map;
    }

    /**
     * 获取不同工作簿的函数式方法
     *
     * @param wb
     * @return
     */
    public static FormulaEvaluator getFormulaEvaluator(Workbook wb) {
        FormulaEvaluator evaluator = null;
        if (wb instanceof XSSFWorkbook) {
            XSSFWorkbook xWb = (XSSFWorkbook) wb;
            evaluator = new XSSFFormulaEvaluator(xWb);
        } else if (wb instanceof HSSFWorkbook) {
            HSSFWorkbook hWb = (HSSFWorkbook) wb;
            evaluator = new HSSFFormulaEvaluator(hWb);
        }
        return evaluator;
    }

    /**
     * 详细转换方法
     *
     * @param wb
     * @return
     * @throws Exception
     */
    private static List<String> getExcelInfo(Workbook wb) throws Exception {
        List<String> list = new ArrayList<String>();
        FormulaEvaluator evaluator = getFormulaEvaluator(wb);
        int sheets = wb.getNumberOfSheets();
        for (int i = 0; i < sheets; i++) {
            list.add(Sheet2Html(wb, evaluator, wb.getSheetAt(i)));
        }
        return list;
    }
/*    private List<String> getExcelInfo(Workbook wb) throws Exception {
        List<String> list=new ArrayList<String>();
        FormulaEvaluator evaluator = getFormulaEvaluator(wb);
        int sheets = wb.getNumberOfSheets();
        for (int i = 0; i < sheets; i++) {
            list.add(Sheet2Html(wb, evaluator, wb.getSheetAt(i)));
        }
        return list;
    }*/

    private static String Sheet2Html(Workbook wb, FormulaEvaluator evaluator, Sheet sheet) throws Exception {
        StringBuffer sb = new StringBuffer();
        int lastRowNum = sheet.getLastRowNum();
        if (lastRowNum > 10000) {
            //funnyZpC_由于数据量大可能会导致内存不足,故限制为1W行内
            throw new Exception("提示：行数过多!");
        }
        Map<String, String> map[] = getRowSpanColSpanMap(sheet);
        //sb.append("<table style='border-collapse:collapse;' >");
        Row row = null;
        Cell cell = null;
        int maxColumn = 0;//funnyZpC 计算最大列数
        for (int rowNum = sheet.getFirstRowNum(); rowNum <= lastRowNum; rowNum++) {
            row = sheet.getRow(rowNum);
            if (row == null) {
                sb.append("<tr><td style='" + getNullCellBorderStyle()
                        + "' > &nbsp;</td></tr>");
                continue;
            }
            sb.append("<tr>");
            int lastColNum = row.getLastCellNum();
            if (lastColNum > maxColumn) {
                maxColumn = lastColNum;//赋值最大列数
            }
            for (int colNum = 0; colNum < lastColNum; colNum++) {
                cell = row.getCell(colNum);
                if (cell == null) {
                    sb.append("<td style='" + getNullCellBorderStyle()
                            + ";white-space: nowrap;'>&nbsp;</td>");
                    continue;
                }

                String stringValue = null;
                // switch (cell.getCellType()) { //获取单元格的值
                // case HSSFCell.CELL_TYPE_BLANK:
                // stringValue = "";
                // break;
                // case HSSFCell.CELL_TYPE_BOOLEAN:
                // stringValue = String
                // .valueOf(cell.getBooleanCellValue());
                // break;
                // case HSSFCell.CELL_TYPE_ERROR:
                // stringValue = cell.getErrorCellString();
                // break;
                // case HSSFCell.CELL_TYPE_FORMULA:
                // stringValue = cell.getCTCell().getV();
                // break;
                // case HSSFCell.CELL_TYPE_NUMERIC:
                // stringValue = String
                // .valueOf(cell.getNumericCellValue());
                // break;
                // case HSSFCell.CELL_TYPE_STRING:
                // stringValue = cell.getStringCellValue();
                // break;
                // default:
                // break;
                // }

                // String stringValue = null;
                // long longVal;
                // double doubleVal;
                // int intvalue;
                switch (cell.getCellType()) {
                    case XSSFCell.CELL_TYPE_NUMERIC: // 数值型
                        if (HSSFDateUtil.isCellDateFormatted(cell)) { // 如果是时间类型
                            SimpleDateFormat sdf = new SimpleDateFormat("yyyy/MM/dd");
                            stringValue = sdf.format(cell.getDateCellValue());
                        } else { // 纯数字
                            double value = cell.getNumericCellValue();
                            CellStyle style1 = cell.getCellStyle();
                            DecimalFormat format = new DecimalFormat();
                            String temp = style1.getDataFormatString();
                            // 单元格设置成常规
                            // if (temp.equals("General")) {
                            if ("General".equals(temp)) {//funnyZpC
                                format.applyPattern("#");
                            }
                            stringValue = format.format(value);
                            // doubleVal = cell.getNumericCellValue();
                            // intvalue = (int) cell.getNumericCellValue();
                            // if(doubleVal == intvalue)
                            // stringValue = String.valueOf(intvalue);
                            // else
                            // // cell.setCellType(HSSFCell.CELL_TYPE_STRING);
                            // // stringValue = cell.getStringCellValue();
                            // stringValue =
                            // String.valueOf(cell.getNumericCellValue());
                        }
                        break;
                    case HSSFCell.CELL_TYPE_STRING: // 字符串型
                        stringValue = cell.getStringCellValue();
                        break;
                    case HSSFCell.CELL_TYPE_BOOLEAN: // 布尔
                        stringValue = " " + cell.getBooleanCellValue();
                        break;
                    case HSSFCell.CELL_TYPE_BLANK: // 空值
                        stringValue = "";
                        break;
                    case HSSFCell.CELL_TYPE_ERROR: // 故障
                        stringValue = "";
                        break;
                    case HSSFCell.CELL_TYPE_FORMULA: // 公式型
                        try {
                            CellValue cellValue;
                            cellValue = evaluator.evaluate(cell);
                            switch (cellValue.getCellType()) { // 判断公式类型
                                case Cell.CELL_TYPE_BOOLEAN:
                                    stringValue = String.valueOf(cellValue
                                            .getBooleanValue());
                                    break;
                                case Cell.CELL_TYPE_NUMERIC:

                                    // 处理日期
                                    if (DateUtil.isCellDateFormatted(cell)) {
                                        SimpleDateFormat sdf = new SimpleDateFormat(
                                                "yyyy/MM/dd");
                                        stringValue = sdf.format(cell
                                                .getDateCellValue());
                                    } else {
                                        // longVal =
                                        // Math.round(cell.getNumericCellValue());
                                        // doubleVal =
                                        // Math.round(cell.getNumericCellValue());
                                        // if(Double.parseDouble(longVal + ".0") ==
                                        // doubleVal)
                                        // stringValue = String.valueOf(longVal);
                                        // else
                                        // stringValue = String.valueOf(doubleVal);
                                        double value = cell.getNumericCellValue();
                                        CellStyle style1 = cell.getCellStyle();
                                        DecimalFormat format = new DecimalFormat();
                                        String temp = style1.getDataFormatString();
                                        // 单元格设置成常规
                                        if (temp.equals("General")) {
                                            format.applyPattern("#");
                                        }
                                        stringValue = format.format(value);
                                    }

                                    break;
                                case Cell.CELL_TYPE_STRING:
                                    stringValue = cellValue.getStringValue();
                                    break;
                                case Cell.CELL_TYPE_BLANK:
                                    stringValue = "";
                                    break;
                                case Cell.CELL_TYPE_ERROR:
                                    stringValue = "";
                                    break;
                                case Cell.CELL_TYPE_FORMULA:
                                    stringValue = "";
                                    break;
                            }
                        } catch (Exception e) {
                            // stringValue = cell.;
                            cell.getCellFormula();
                        }
                        break;
                    default:
                        stringValue = cell.getStringCellValue().toString();
                        break;
                }

                // switch (cell.getCellType()) {
                // case HSSFCell.CELL_TYPE_FORMULA:
                // // cell.getCellFormula();
                // try {
                // stringValue = String.valueOf(cell.getNumericCellValue());
                // } catch (IllegalStateException e) {
                // stringValue =
                // String.valueOf(cell.getRichStringCellValue());
                // }
                // break;
                // case HSSFCell.CELL_TYPE_NUMERIC:
                // stringValue = String.valueOf(cell.getNumericCellValue());
                // break;
                // case HSSFCell.CELL_TYPE_STRING:
                // stringValue =
                // String.valueOf(cell.getRichStringCellValue());
                // break;
                // }
                //
                if (map[0].containsKey(rowNum + "," + colNum)) {
                    String pointString = map[0].get(rowNum + "," + colNum);
                    map[0].remove(rowNum + "," + colNum);
                    int bottomeRow = Integer
                            .valueOf(pointString.split(",")[0]);
                    int bottomeCol = Integer
                            .valueOf(pointString.split(",")[1]);
                    int rowSpan = bottomeRow - rowNum + 1;
                    int colSpan = bottomeCol - colNum + 1;
                    sb.append("<td  rowspan= '" + rowSpan + "' colspan= '"
                            + colSpan + "' ");

                } else if (map[1].containsKey(rowNum + "," + colNum)) {
                    map[1].remove(rowNum + "," + colNum);
                    continue;
                } else {
                    sb.append("<td ");
                }
                // 获取样式的内容
                if (wb instanceof XSSFWorkbook) {
                    XSSFCellStyle xcellStyle = ((XSSFCell) cell).getCellStyle();
                    if (xcellStyle != null) {
                        short alignment = xcellStyle.getAlignment();
                        sb.append("align='" + convertAlignToHtml(alignment)
                                + "' ");
                        short verticalAlignment = xcellStyle
                                .getVerticalAlignment();
                        sb.append("valign='"
                                + convertVerticalAlignToHtml(verticalAlignment)
                                + "' ");
                        sb.append("style='");
                        XSSFFont xf = xcellStyle.getFont();
                        short boldWeight = xf.getBoldweight();
                        XSSFColor xc = xf.getXSSFColor();
                        String fontColorStr = ColorUtil
                                .convertColorToHex(xc);
                        String fontName = xf.getFontName();
                        int fontsize = xf.getFontHeightInPoints();
                        int columnWidth = (int) sheet.getColumnWidthInPixels(cell
                                .getColumnIndex());
                        int rowHeight = (int) row.getHeightInPoints();
                        sb.append("width:" + columnWidth + "px;");
                        sb.append("height:" + rowHeight + "px;");
                        if (fontColorStr != null
                                && !"".equals(fontColorStr.trim())) {
                            sb.append("color:" + fontColorStr + ";"); // 字体颜色
                        }
                        if (fontName != null && !"".equals(fontName.trim())) {
                            sb.append("font-family:\"" + fontName + "\";"); // 字体
                        }
                        if (fontsize != 0) {
                            sb.append("font-size:" + fontsize + "px;"); // 字体大小
                        }
                        XSSFColor xbgColor = null;
                        if (xcellStyle.getFillPattern() == CellStyle.SOLID_FOREGROUND) {
                            xbgColor = xcellStyle
                                    .getFillForegroundXSSFColor();
                        }
                        xbgColor = xcellStyle.getFillForegroundXSSFColor();
                        String bgColorStr = ColorUtil
                                .convertColorToHex(xbgColor);
                        if (bgColorStr != null && !"".equals(bgColorStr.trim())) {
                            sb.append("background-color:" + bgColorStr + ";"); // 背景颜色
                        }
                        sb.append(getBorderStyle(0, xcellStyle.getBorderTop(), xcellStyle.getTopBorderXSSFColor()));
                        sb.append(getBorderStyle(1, xcellStyle.getBorderRight(), xcellStyle.getRightBorderXSSFColor()));
                        sb.append(getBorderStyle(2, xcellStyle.getBorderBottom(), xcellStyle.getBottomBorderXSSFColor()));
                        sb.append(getBorderStyle(3, xcellStyle.getBorderLeft(), xcellStyle.getLeftBorderXSSFColor()));
                        sb.append("font-weight:" + boldWeight + ";"); // 字体加粗
                        sb.append("font-size: " + xf.getFontHeight() / 2.5 + "%;"); // 字体大小
                        sb.append("white-space: nowrap;");
                    }
                } else if (wb instanceof HSSFWorkbook) {

                    HSSFCellStyle hcellStyle = ((HSSFCell) cell)
                            .getCellStyle();
                    if (hcellStyle != null) {
                        short alignment = hcellStyle.getAlignment();
                        sb.append("align='" + convertAlignToHtml(alignment)
                                + "' ");
                        short verticalAlignment = hcellStyle
                                .getVerticalAlignment();
                        sb.append("valign='"
                                + convertVerticalAlignToHtml(verticalAlignment)
                                + "' ");
                        sb.append("style='");
                        HSSFFont hf = hcellStyle.getFont(wb);
                        short boldWeight = hf.getBoldweight();
                        short fontColor = hf.getColor();
                        String fontName = hf.getFontName();
                        int fontsize = hf.getFontHeightInPoints();
                        HSSFPalette palette = ((HSSFWorkbook) wb)
                                .getCustomPalette(); // 类HSSFPalette用于求的颜色的国际标准形式
                        HSSFColor hc = palette.getColor(fontColor);
                        String fontColorStr = ColorUtil
                                .convertColorToHex(hc);
                        int columnWidth = (int) sheet.getColumnWidthInPixels(cell
                                .getColumnIndex());
                        int rowHeight = (int) row.getHeightInPoints();
                        sb.append("width:" + columnWidth + "px;");
                        sb.append("height:" + rowHeight + "px;");
                        if (fontColorStr != null
                                && !"".equals(fontColorStr.trim())) {
                            sb.append("color:" + fontColorStr + ";"); // 字体颜色
                        }
                        if (fontName != null && !"".equals(fontName.trim())) {
                            sb.append("font-family:\"" + fontName + "\";"); // 字体
                        }
                        if (fontsize != 0) {
                            sb.append("font-size:" + fontsize + "px;"); // 字体大小
                        }
                        short bgColor = hcellStyle.getFillForegroundColor();
                        hc = palette.getColor(bgColor);
                        String bgColorStr = convertToStardColor(hc);
                        if (bgColorStr != null
                                && !"".equals(bgColorStr.trim())) {
                            sb.append("background-color:" + bgColorStr
                                    + ";"); // 背景颜色
                        }
                        sb.append(getBorderStyle(palette, 0,
                                hcellStyle.getBorderTop(),
                                hcellStyle.getTopBorderColor()));
                        sb.append(getBorderStyle(palette, 1,
                                hcellStyle.getBorderRight(),
                                hcellStyle.getRightBorderColor()));
                        sb.append(getBorderStyle(palette, 3,
                                hcellStyle.getBorderLeft(),
                                hcellStyle.getLeftBorderColor()));
                        sb.append(getBorderStyle(palette, 2,
                                hcellStyle.getBorderBottom(),
                                hcellStyle.getBottomBorderColor()));
                        sb.append("font-weight:" + boldWeight + ";"); // 字体加粗
                        sb.append("font-size: " + hf.getFontHeight() / 2.5
                                + "%;"); // 字体大小
                        sb.append("white-space: nowrap;");
                    }
                }
                sb.append("' ");
                sb.append(">");
                if (stringValue == null || "".equals(stringValue.trim())) {
                    sb.append(" &nbsp; ");
                } else {
                    // 将ascii码为160的空格转换为html下的空格（&nbsp;）
                    sb.append(stringValue.replace(
                            String.valueOf((char) 160), "&nbsp;"));
                }
                sb.append("</td>");
            }
            sb.append("</tr>");
        }
        String head = String.format("%s%s%s", "<table style='border-collapse:collapse;margin:1%' ><thead><tr><th colspan=\"", maxColumn, "\">" + sheet.getSheetName() + "</th></tr></thead>");
        sb.append("</table>");
        return String.format("%s%s", head, sb.toString());
    }

    /**
     * 转换excel2html方法
     *
     * @param wb 工作簿
     * @return map key:sheet1 value:
     * <table>
     * ...
     * </table>
     * 字符串
     */
    public static List<String> getExcelToHtml(Workbook wb) {
        try {
            List<String> htmlPage = getExcelInfo(wb);
            return htmlPage;
        } catch (Exception e) {
            e.printStackTrace();
        }
        return null;
    }

    /**
     * 程序入口方法
     *
     * @param filePath 文件的路径
     * @return <table>
     * ...
     * </table>
     * 字符串
     */
    public static List<String> readExcelToHtml(String filePath) {
        List<String> htmlExcel = null;
        try {
            File sourcefile = new File(filePath);
            InputStream is = new FileInputStream(sourcefile);
            Workbook wb = WorkbookFactory.create(is);
            htmlExcel = getExcelToHtml(wb);
        } catch (EncryptedDocumentException e) {
            e.printStackTrace();
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (InvalidFormatException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
        return htmlExcel;

    }
}
