package com.funny.excelbuddy.export.utils;

import com.google.common.collect.Maps;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellRangeAddressBase;
import org.apache.poi.util.Units;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.*;
import org.openxmlformats.schemas.drawingml.x2006.main.*;
import org.openxmlformats.schemas.drawingml.x2006.spreadsheetDrawing.CTConnector;
import org.openxmlformats.schemas.drawingml.x2006.spreadsheetDrawing.CTShape;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTAutoFilter;
import org.springframework.core.io.ClassPathResource;
import org.springframework.core.io.Resource;
import org.springframework.util.CollectionUtils;

import java.io.IOException;
import java.io.InputStream;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Optional;
import java.util.function.Consumer;
import java.util.stream.Collectors;

public class ExcelUtil {

    /**
     * 內存保留行數
     */
    public static final int ROW_ACCESS_WINDOW_SIZE = 1000;

    /**
     * 初始化工作簿并创建第一页
     *
     * @return
     */
    public static SXSSFWorkbook initWorkbook() {
        SXSSFWorkbook resultWB = new SXSSFWorkbook(ROW_ACCESS_WINDOW_SIZE);
        resultWB.createSheet("导出数据");
        return resultWB;
    }

    /**
     * 初始化表头
     *
     * @param sheet   Excel工具簿
     * @param headers 表头数据
     */
    public static void fillRowHeader(SXSSFSheet sheet, String[] headers) {
        SXSSFRow row = sheet.createRow(0);
        //填充表头
        for (int i = 0; i < headers.length; i++) {
            row.createCell(i).setCellValue(headers[i]);
        }
    }

    /**
     * 填充数据行
     *
     * @param sheet
     * @param rowIndex
     * @param rowData
     */
    public static void fillRowData(SXSSFSheet sheet, int rowIndex, Object[] rowData) {
        SXSSFRow row = sheet.createRow(rowIndex);
        for (int i = 0; i < rowData.length; i++) {
            row.createCell(rowIndex).setCellValue(Optional.ofNullable(rowData[i]).map(t -> t.toString()).orElse(""));
        }
    }

    /**
     * 填充数据行
     *
     * @param sheet
     * @param rowIndex
     * @param consumer
     */
    public static void fillRowData(SXSSFSheet sheet, int rowIndex, Consumer<SXSSFRow> consumer) {
        SXSSFRow row = sheet.createRow(rowIndex);
        consumer.accept(row);
    }

    /**
     * 初始化工作薄
     *
     * @param template
     * @return
     */
    public static SXSSFWorkbook initWorkbook(String template) {
        SXSSFWorkbook resultWB = null;
        try {
//            InputStream ExcelFileToRead = new FileInputStream(template);
            Resource resource = new ClassPathResource(template);
            InputStream is = resource.getInputStream();
            XSSFWorkbook wb = new XSSFWorkbook(is);
            //内存只留1000行
            resultWB = new SXSSFWorkbook(ROW_ACCESS_WINDOW_SIZE);
            //从模板xlsx复制到目标文件
            ExcelUtil.cloneWorkbook(wb, resultWB);
        } catch (IOException e) {
            e.printStackTrace();
        }
        return resultWB;
    }

    public static Row getRow(SXSSFWorkbook resultWB, int rowIndex) {
        SXSSFSheet sheet = resultWB.getSheetAt(0);
        return sheet.getRow(rowIndex);
    }

    public static Row getSrcRow(SXSSFWorkbook resultWB, int copyRowIndex) {
        return getSrcRow(resultWB, 0, copyRowIndex);
    }

    public static Row getSrcRow(SXSSFWorkbook resultWB, int sheetIndex, int copyRowIndex) {
        SXSSFSheet sheet = resultWB.getSheetAt(sheetIndex);
        return sheet.getRow(copyRowIndex);
    }

    /**
     * 获取excel行，
     *
     * @param resultWB
     * @param srcRow   复制格式的行
     * @param rowIndex 行索引
     * @return
     */
    public static Row getDestRow(SXSSFWorkbook resultWB, Row srcRow, int rowIndex) {
        return getDestRow(resultWB, srcRow, 0, rowIndex);
    }

    /**
     * 获取excel行，
     *
     * @param resultWB
     * @param srcRow   复制格式的行
     * @param rowIndex 行索引
     * @return
     */
    public static Row getDestRow(SXSSFWorkbook resultWB, Row srcRow, int sheetIndex, int rowIndex) {
        SXSSFSheet sheet = resultWB.getSheetAt(sheetIndex);
        Map<Integer, CellStyle> styleMap = Maps.newHashMap();
        Row destRow = sheet.createRow(rowIndex);
        ExcelUtil.copyRow(srcRow, destRow, styleMap);
        return destRow;
    }

    public static void cloneWorkbook(XSSFWorkbook fromWB, SXSSFWorkbook resultWB) {
        for (int i = 0; i < fromWB.getNumberOfSheets(); i++) {
            XSSFSheet sheet = fromWB.getSheetAt(i);
            SXSSFSheet newSheet = resultWB.createSheet(sheet.getSheetName());
            cloneSheet(sheet, newSheet);
        }
    }

    public static void cloneSheet(XSSFSheet sheet, SXSSFSheet newSheet) {
        CTAutoFilter autoFilter = sheet.getCTWorksheet().getAutoFilter();
        if (autoFilter != null) {
            newSheet.setAutoFilter(CellRangeAddress.valueOf(autoFilter.getRef()));
        }

        int maxColumnNum = 0;
        Map<Integer, CellStyle> styleMap = new HashMap<Integer, CellStyle>();
        for (int i = sheet.getFirstRowNum(); i <= sheet.getLastRowNum(); i++) {
            Row srcRow = sheet.getRow(i);
            Row destRow = newSheet.createRow(i);
            if (srcRow != null) {
                copyRow(srcRow, destRow, styleMap);
                if (srcRow.getLastCellNum() > maxColumnNum) {
                    maxColumnNum = srcRow.getLastCellNum();
                }
            }
        }
        for (int i = 0; i <= maxColumnNum; i++) {
            newSheet.setColumnWidth(i, sheet.getColumnWidth(i));
        }
        //增加复制合并单元格
        copyMergedRegions(sheet, newSheet);
    }

    /**
     * 刷新
     *
     * @param resultWB
     */
    public static void flushRow(SXSSFWorkbook resultWB) {
        try {
            SXSSFSheet sheet = resultWB.getSheetAt(0);
            sheet.flushRows();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    public static void copyRow(Row srcRow, Row destRow, Map<Integer, CellStyle> styleMap) {
        destRow.setHeight(srcRow.getHeight());
        // reckoning delta rows
        int deltaRows = destRow.getRowNum() - srcRow.getRowNum();
        // pour chaque row
        for (int j = srcRow.getFirstCellNum(); j <= srcRow.getLastCellNum(); j++) {
            Cell oldCell = srcRow.getCell(j);   // ancienne cell
            Cell newCell = destRow.getCell(j);  // new cell
            if (oldCell != null) {
                if (newCell == null) {
                    newCell = destRow.createCell(j);
                }
                // copy chaque cell
                copyCell(oldCell, newCell, styleMap);
                // copy les informations de fusion entre les cellules

            }
        }
    }

    public static void copyCell(Cell oldCell, Cell newCell, Map<Integer, CellStyle> styleMap) {
        if (styleMap != null) {
            if (oldCell.getSheet().getWorkbook() == newCell.getSheet().getWorkbook()) {
                newCell.setCellStyle(oldCell.getCellStyle());
            } else {
                int stHashCode = oldCell.getCellStyle().hashCode();
                CellStyle newCellStyle = styleMap.get(stHashCode);
                if (newCellStyle == null) {
                    newCellStyle = newCell.getSheet().getWorkbook().createCellStyle();
                    newCellStyle.cloneStyleFrom(oldCell.getCellStyle());
                    styleMap.put(stHashCode, newCellStyle);
                }
                newCell.setCellStyle(newCellStyle);
            }
        }
//        CellStyle newCellStyle = newCell.getSheet().getWorkbook().createCellStyle();
//        copyCellStyle(oldCell.getCellStyle(), newCellStyle, oldCell.getSheet().getWorkbook());

        switch (oldCell.getCellType()) {
            case STRING:
                newCell.setCellValue(oldCell.getStringCellValue());
                break;
            case NUMERIC:
                newCell.setCellValue(oldCell.getNumericCellValue());
                break;
            case BLANK:
                newCell.setBlank();
                break;
            case BOOLEAN:
                newCell.setCellValue(oldCell.getBooleanCellValue());
                break;
            case ERROR:
                newCell.setCellErrorValue(oldCell.getErrorCellValue());
                break;
            case FORMULA:
                newCell.setCellFormula(oldCell.getCellFormula());
                break;
            default:
                break;
        }

    }

    public static void copyCellStyle(CellStyle fromStyle, CellStyle toStyle, Workbook workbook) {
        //水平垂直对齐方式
        toStyle.setAlignment(fromStyle.getAlignmentEnum());
        toStyle.setVerticalAlignment(fromStyle.getVerticalAlignmentEnum());
        //边框和边框颜色
        toStyle.setBorderBottom(fromStyle.getBorderBottomEnum());
        toStyle.setBorderLeft(fromStyle.getBorderLeftEnum());
        toStyle.setBorderRight(fromStyle.getBorderRightEnum());
        toStyle.setBorderTop(fromStyle.getBorderTopEnum());
        toStyle.setTopBorderColor(fromStyle.getTopBorderColor());
        toStyle.setBottomBorderColor(fromStyle.getBottomBorderColor());
        toStyle.setRightBorderColor(fromStyle.getRightBorderColor());
        toStyle.setLeftBorderColor(fromStyle.getLeftBorderColor());
        //背景和前景
        if (fromStyle instanceof XSSFCellStyle) {
            XSSFCellStyle xssfToStyle = (XSSFCellStyle) toStyle;
            xssfToStyle.setFillBackgroundColor(((XSSFCellStyle) fromStyle).getFillBackgroundColorColor());
            xssfToStyle.setFillForegroundColor(((XSSFCellStyle) fromStyle).getFillForegroundColorColor());
        } else {
            toStyle.setFillBackgroundColor(fromStyle.getFillBackgroundColor());
            toStyle.setFillForegroundColor(fromStyle.getFillForegroundColor());
        }
        toStyle.setDataFormat(fromStyle.getDataFormat());
        toStyle.setFillPattern(fromStyle.getFillPatternEnum());
        if (fromStyle instanceof XSSFCellStyle) {
            toStyle.setFont(((XSSFCellStyle) fromStyle).getFont());
        } else if (fromStyle instanceof HSSFCellStyle) {
            toStyle.setFont(((HSSFCellStyle) fromStyle).getFont(workbook));
        }
        toStyle.setHidden(fromStyle.getHidden());
        //首行缩进
        toStyle.setIndention(fromStyle.getIndention());
        toStyle.setLocked(fromStyle.getLocked());
        //旋转
        toStyle.setRotation(fromStyle.getRotation());
        toStyle.setWrapText(fromStyle.getWrapText());
    }

    /**
     * 复制合并单元格
     *
     * @param sheet
     * @param newSheet
     */
    public static void copyMergedRegions(XSSFSheet sheet, SXSSFSheet newSheet) {
        //获取合并单元格
        List<CellRangeAddress> cellRangeAddressList = sheet.getMergedRegions();
        if (CollectionUtils.isEmpty(cellRangeAddressList)) {
            return;
        }
        Map<Integer, List<CellRangeAddress>> rowCellRangeAddressMap = cellRangeAddressList.stream().collect(Collectors.groupingBy(CellRangeAddressBase::getFirstRow));

        for (int i = sheet.getFirstRowNum(); i <= sheet.getLastRowNum(); i++) {
            Row toRow = newSheet.getRow(i);
            //复制合并单元格
            List<CellRangeAddress> rowCellRangeAddressList = rowCellRangeAddressMap.get(i);
            if (CollectionUtils.isEmpty(rowCellRangeAddressList)) {
                //当前行没有合并的单元格，不处理
                return;
            }
            for (CellRangeAddress cellRangeAddress : rowCellRangeAddressList) {
                CellRangeAddress newCellRangeAddress = new CellRangeAddress(toRow.getRowNum(), (toRow.getRowNum() +
                        (cellRangeAddress.getLastRow() - cellRangeAddress.getFirstRow())), cellRangeAddress
                        .getFirstColumn(), cellRangeAddress.getLastColumn());
                newSheet.addMergedRegionUnsafe(newCellRangeAddress);
            }
        }
    }

    /**
     * 复制XSSFSimpleShape类
     *
     * @param fromShape
     * @param drawing
     * @param anchor
     * @return
     */
    public static XSSFSimpleShape copyXSSFSimpleShape(XSSFSimpleShape fromShape, XSSFDrawing drawing, XSSFClientAnchor anchor) {
        XSSFSimpleShape toShape = drawing.createSimpleShape(anchor);
        CTShape ctShape = fromShape.getCTShape();
        CTShapeProperties ctShapeProperties = ctShape.getSpPr();
        CTLineProperties lineProperties = ctShapeProperties.isSetLn() ? ctShapeProperties.getLn() : ctShapeProperties.addNewLn();
        CTPresetLineDashProperties dashStyle = lineProperties.isSetPrstDash() ? lineProperties.getPrstDash() : CTPresetLineDashProperties.Factory.newInstance();
        STPresetLineDashVal.Enum dashStyleEnum = dashStyle.isSetVal() ? dashStyle.getVal() : STPresetLineDashVal.Enum.forInt(1);
        CTSolidColorFillProperties fill = lineProperties.isSetSolidFill() ? lineProperties.getSolidFill() : lineProperties.addNewSolidFill();
        CTSRgbColor rgb = fill.isSetSrgbClr() ? fill.getSrgbClr() : CTSRgbColor.Factory.newInstance();
        // 设置形状类型
        toShape.setShapeType(fromShape.getShapeType());
        // 设置线宽
        toShape.setLineWidth(lineProperties.getW() * 1.0 / Units.EMU_PER_POINT);
        // 设置线的风格
        toShape.setLineStyle(dashStyleEnum.intValue() - 1);
        // 设置线的颜色
        byte[] rgbBytes = rgb.getVal();
        if (rgbBytes == null) {
            toShape.setLineStyleColor(0, 0, 0);
        } else {
            toShape.setLineStyleColor(rgbBytes[0], rgbBytes[1], rgbBytes[2]);
        }
        return toShape;
    }

    /**
     * 复制XSSFConnector类
     *
     * @param fromShape
     * @param drawing
     * @param anchor
     * @return
     */
    public static XSSFConnector copyXSSFConnector(XSSFConnector fromShape, XSSFDrawing drawing, XSSFClientAnchor anchor) {
        XSSFConnector toShape = drawing.createConnector(anchor);
        CTConnector ctConnector = fromShape.getCTConnector();
        CTShapeProperties ctShapeProperties = ctConnector.getSpPr();
        CTLineProperties lineProperties = ctShapeProperties.isSetLn() ? ctShapeProperties.getLn() : ctShapeProperties.addNewLn();
        CTPresetLineDashProperties dashStyle = lineProperties.isSetPrstDash() ? lineProperties.getPrstDash() : CTPresetLineDashProperties.Factory.newInstance();
        STPresetLineDashVal.Enum dashStyleEnum = dashStyle.isSetVal() ? dashStyle.getVal() : STPresetLineDashVal.Enum.forInt(1);
        CTSolidColorFillProperties fill = lineProperties.isSetSolidFill() ? lineProperties.getSolidFill() : lineProperties.addNewSolidFill();
        CTSRgbColor rgb = fill.isSetSrgbClr() ? fill.getSrgbClr() : CTSRgbColor.Factory.newInstance();
        // 设置形状类型
        toShape.setShapeType(fromShape.getShapeType());
        // 设置线宽
        toShape.setLineWidth(lineProperties.getW() * 1.0 / Units.EMU_PER_POINT);
        // 设置线的风格
        toShape.setLineStyle(dashStyleEnum.intValue() - 1);
        // 设置线的颜色
        byte[] rgbBytes = rgb.getVal();
        if (rgbBytes == null) {
            toShape.setLineStyleColor(0, 0, 0);
        } else {
            toShape.setLineStyleColor(rgbBytes[0], rgbBytes[1], rgbBytes[2]);
        }
        return toShape;

    }
}