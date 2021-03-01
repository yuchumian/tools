package com.yuchumian.tools.excel;

import lombok.extern.slf4j.Slf4j;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.*;

import java.util.HashMap;

/**
 * @author yuchumian 2021-02-26
 **/
@Slf4j
public class ExcelTransform {
    private int lastColumn = 0;
    private final HashMap<Integer, HSSFCellStyle> styleMap = new HashMap<>();

    public void transformXlsx(XSSFWorkbook xlsxWorkbook, HSSFWorkbook xlsWorkbook) {
        HSSFSheet xlsSheet;
        XSSFSheet xlsxSheet;
        xlsWorkbook.setMissingCellPolicy(xlsxWorkbook.getMissingCellPolicy());
        for (int i = 0; i < xlsxWorkbook.getNumberOfSheets(); i++) {
            xlsxSheet = xlsxWorkbook.getSheetAt(i);
            xlsSheet = xlsWorkbook.createSheet(xlsxSheet.getSheetName());
            this.transform(xlsxWorkbook, xlsWorkbook, xlsxSheet, xlsSheet);
        }
    }

    private void transform(XSSFWorkbook xlsxWorkbook, HSSFWorkbook xlsWorkbook,
                           XSSFSheet xlsxSheet, HSSFSheet xlsSheet) {

        xlsSheet.setDisplayFormulas(xlsxSheet.isDisplayFormulas());
        xlsSheet.setDisplayGridlines(xlsxSheet.isDisplayGridlines());
        xlsSheet.setDisplayGuts(xlsxSheet.getDisplayGuts());
        xlsSheet.setDisplayRowColHeadings(xlsxSheet.isDisplayRowColHeadings());
        xlsSheet.setDisplayZeros(xlsxSheet.isDisplayZeros());
        xlsSheet.setFitToPage(xlsxSheet.getFitToPage());

        xlsSheet.setHorizontallyCenter(xlsxSheet.getHorizontallyCenter());
        xlsSheet.setMargin(Sheet.BottomMargin,
                xlsxSheet.getMargin(Sheet.BottomMargin));
        xlsSheet.setMargin(Sheet.FooterMargin,
                xlsxSheet.getMargin(Sheet.FooterMargin));
        xlsSheet.setMargin(Sheet.HeaderMargin,
                xlsxSheet.getMargin(Sheet.HeaderMargin));
        xlsSheet.setMargin(Sheet.LeftMargin,
                xlsxSheet.getMargin(Sheet.LeftMargin));
        xlsSheet.setMargin(Sheet.RightMargin,
                xlsxSheet.getMargin(Sheet.RightMargin));
        xlsSheet.setMargin(Sheet.TopMargin, xlsxSheet.getMargin(Sheet.TopMargin));
        xlsSheet.setPrintGridlines(xlsSheet.isPrintGridlines());
        xlsSheet.setRightToLeft(xlsSheet.isRightToLeft());
        xlsSheet.setRowSumsBelow(xlsSheet.getRowSumsBelow());
        xlsSheet.setRowSumsRight(xlsSheet.getRowSumsRight());
        xlsSheet.setVerticallyCenter(xlsxSheet.getVerticallyCenter());

        HSSFRow xlsRow;
        for (Row row : xlsxSheet) {
            xlsRow = xlsSheet.createRow(row.getRowNum());
            this.transform(xlsxWorkbook, xlsWorkbook, (XSSFRow) row, xlsRow);
        }

        for (int i = 0; i < this.lastColumn; i++) {
            xlsSheet.setColumnWidth(i, xlsxSheet.getColumnWidth(i));
            xlsSheet.setColumnHidden(i, xlsxSheet.isColumnHidden(i));
        }

        for (int i = 0; i < xlsxSheet.getNumMergedRegions(); i++) {
            CellRangeAddress merged = xlsxSheet.getMergedRegion(i);
            xlsSheet.addMergedRegion(merged);
        }
    }

    private void transform(XSSFWorkbook xlsxWorkbook, HSSFWorkbook xlsWorkbook,
                           XSSFRow xlsxRow, HSSFRow xlsRow) {
        HSSFCell cellXls;
        xlsRow.setHeight(xlsxRow.getHeight());

        for (Cell cell : xlsxRow) {
            cellXls = xlsRow.createCell(cell.getColumnIndex(),
                    cell.getCellTypeEnum());
            this.transform(xlsxWorkbook, xlsWorkbook, (XSSFCell) cell,
                    cellXls);
        }
        this.lastColumn = Math.max(this.lastColumn, xlsxRow.getLastCellNum());
    }

    private void transform(XSSFWorkbook xlsxWorkbook, HSSFWorkbook xlsWorkbook,
                           XSSFCell cellXlsx, HSSFCell cellXls) {
        cellXls.setCellComment(cellXlsx.getCellComment());

        Integer hash = cellXlsx.getCellStyle().hashCode();
        if (!this.styleMap.containsKey(hash)) {
            this.transform(xlsxWorkbook, xlsWorkbook, hash,
                    cellXlsx.getCellStyle(),
                    xlsWorkbook.createCellStyle());
        }
        cellXls.setCellStyle(this.styleMap.get(hash));

        switch (cellXlsx.getCellTypeEnum()) {
            case BLANK:
                break;
            case BOOLEAN:
                cellXls.setCellValue(cellXlsx.getBooleanCellValue());
                break;
            case ERROR:
                cellXls.setCellValue(cellXlsx.getErrorCellValue());
                break;
            case FORMULA:
                cellXls.setCellValue(cellXlsx.getCellFormula());
                break;
            case NUMERIC:
                cellXls.setCellValue(cellXlsx.getNumericCellValue());
                break;
            case STRING:
                cellXls.setCellValue(cellXlsx.getStringCellValue());
                break;
            default:
                log.warn("transform: unknown CellType {}", cellXlsx.getCellTypeEnum());
                break;
        }
    }

    private void transform(XSSFWorkbook xlsxWorkbook, HSSFWorkbook xlsWorkbook,
                           Integer hash, XSSFCellStyle styleXlsx, HSSFCellStyle styleXls) {
        styleXls.setAlignment(styleXlsx.getAlignmentEnum());
        styleXls.setBorderBottom(styleXlsx.getBorderBottomEnum());
        styleXls.setBorderLeft(styleXlsx.getBorderLeftEnum());
        styleXls.setBorderRight(styleXlsx.getBorderRightEnum());
        styleXls.setBorderTop(styleXlsx.getBorderTopEnum());
        styleXls.setDataFormat(this.transform(xlsxWorkbook, xlsWorkbook,
                styleXlsx.getDataFormat()));
        styleXls.setFillBackgroundColor(styleXlsx.getFillBackgroundColor());
        styleXls.setFillForegroundColor(styleXlsx.getFillForegroundColor());
        styleXls.setFillPattern(styleXlsx.getFillPatternEnum());
        styleXls.setFont(this.transform(xlsWorkbook,
                styleXlsx.getFont()));
        styleXls.setHidden(styleXlsx.getHidden());
        styleXls.setIndention(styleXlsx.getIndention());
        styleXls.setLocked(styleXlsx.getLocked());
        styleXls.setVerticalAlignment(styleXlsx.getVerticalAlignmentEnum());
        styleXls.setWrapText(styleXlsx.getWrapText());
        this.styleMap.put(hash, styleXls);
    }

    private short transform(XSSFWorkbook xlsxWorkbook, HSSFWorkbook xlsWorkbook,
                            short index) {
        DataFormat xlsxFormat = xlsxWorkbook.createDataFormat();
        DataFormat xlsFormat = xlsWorkbook.createDataFormat();
        return xlsFormat.getFormat(xlsxFormat.getFormat(index));
    }

    /**
     * 转换字体
     * @param xlsWorkbook HSSFWorkbook xls
     * @param xlsxFont  XSSFFont xlsx
     * @return HSSFFont xls
     */
    private HSSFFont transform(HSSFWorkbook xlsWorkbook, XSSFFont xlsxFont) {
        HSSFFont xlsFont = xlsWorkbook.createFont();
        xlsFont.setBold(xlsxFont.getBold());
        xlsFont.setCharSet(xlsxFont.getCharSet());
        xlsFont.setColor(xlsxFont.getColor());
        xlsFont.setFontName(xlsxFont.getFontName());
        xlsFont.setFontHeight(xlsxFont.getFontHeight());
        xlsFont.setItalic(xlsxFont.getItalic());
        xlsFont.setStrikeout(xlsxFont.getStrikeout());
        xlsFont.setTypeOffset(xlsxFont.getTypeOffset());
        xlsFont.setUnderline(xlsxFont.getUnderline());
        return xlsFont;
    }
}
