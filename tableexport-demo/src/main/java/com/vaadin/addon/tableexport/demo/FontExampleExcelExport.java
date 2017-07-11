package com.vaadin.addon.tableexport.demo;

import com.vaadin.addon.tableexport.ExcelExport;
import com.vaadin.addon.tableexport.TableHolder;
import com.vaadin.ui.Table;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.util.HSSFColor.HSSFColorPredefined;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;

public class FontExampleExcelExport extends ExcelExport {

    private static final long serialVersionUID = 3717947558186318581L;

    public FontExampleExcelExport(final TableHolder tableHolder, final String sheetName) {
        super(tableHolder, sheetName);
        format();
    }

    public FontExampleExcelExport(final Table table, final String sheetName) {
        super(table, sheetName);
        format();
    }

    private void format() {
        this.setRowHeaders(true);
        CellStyle style;
        Font f;

        style = this.getTitleStyle();
        style.setFillForegroundColor(HSSFColorPredefined.GREY_25_PERCENT.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND.getCode());
        f = workbook.createFont();
        f.setFontHeightInPoints((short) 12);
        f.setFontName(HSSFFont.FONT_ARIAL);
        f.setColor(HSSFColorPredefined.BLACK.getIndex());
        f.setBold(true);
        style.setFont(f);
        style.setAlignment(HorizontalAlignment.CENTER_SELECTION.getCode());
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderBottom(BorderStyle.THIN);
        style.setLeftBorderColor(HSSFColorPredefined.BLACK.getIndex());
        style.setRightBorderColor(HSSFColorPredefined.BLACK.getIndex());
        style.setTopBorderColor(HSSFColorPredefined.BLACK.getIndex());
        style.setBottomBorderColor(HSSFColorPredefined.BLACK.getIndex());

        style = this.getColumnHeaderStyle();
        style.setFillForegroundColor(HSSFColorPredefined.GREY_25_PERCENT.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND.getCode());
        f = workbook.createFont();
        f.setFontHeightInPoints((short) 12);
        f.setFontName(HSSFFont.FONT_ARIAL);
        f.setColor(HSSFColorPredefined.BLACK.getIndex());
        f.setBold(true);
        style.setFont(f);
        style.setAlignment(HorizontalAlignment.CENTER.getCode());
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderBottom(BorderStyle.THIN);
        style.setLeftBorderColor(HSSFColorPredefined.BLACK.getIndex());
        style.setRightBorderColor(HSSFColorPredefined.BLACK.getIndex());
        style.setTopBorderColor(HSSFColorPredefined.BLACK.getIndex());
        style.setBottomBorderColor(HSSFColorPredefined.BLACK.getIndex());

        style = this.getTotalsDoubleStyle();
        style.setFillForegroundColor(HSSFColorPredefined.WHITE.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND.getCode());
        f = workbook.createFont();
        f.setFontHeightInPoints((short) 12);
        f.setFontName(HSSFFont.FONT_ARIAL);
        f.setColor(HSSFColorPredefined.BLACK.getIndex());
        f.setBold(true);
        style.setFont(f);
        style.setAlignment(HorizontalAlignment.RIGHT.getCode());
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderBottom(BorderStyle.THIN);
        style.setLeftBorderColor(HSSFColorPredefined.BLACK.getIndex());
        style.setRightBorderColor(HSSFColorPredefined.BLACK.getIndex());
        style.setTopBorderColor(HSSFColorPredefined.BLACK.getIndex());
        style.setBottomBorderColor(HSSFColorPredefined.BLACK.getIndex());

        style = this.getDoubleDataStyle();
        style.setFillForegroundColor(HSSFColorPredefined.WHITE.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND.getCode());
        f = workbook.getFontAt(style.getFontIndex());
        f.setFontHeightInPoints((short) 12);
        f.setFontName(HSSFFont.FONT_ARIAL);
        f.setColor(HSSFColorPredefined.BLACK.getIndex());
        f.setBold(false);
        style.setFont(f);
        style.setAlignment(HorizontalAlignment.RIGHT.getCode());
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderBottom(BorderStyle.THIN);
        style.setLeftBorderColor(HSSFColorPredefined.BLACK.getIndex());
        style.setRightBorderColor(HSSFColorPredefined.BLACK.getIndex());
        style.setTopBorderColor(HSSFColorPredefined.BLACK.getIndex());
        style.setBottomBorderColor(HSSFColorPredefined.BLACK.getIndex());

        style = this.getIntegerDataStyle();
        style.setFillForegroundColor(HSSFColorPredefined.WHITE.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND.getCode());
        f = workbook.getFontAt(style.getFontIndex());
        f.setFontHeightInPoints((short) 12);
        f.setFontName(HSSFFont.FONT_ARIAL);
        f.setColor(HSSFColorPredefined.BLACK.getIndex());
        f.setBold(false);
        style.setFont(f);
        style.setAlignment(HorizontalAlignment.RIGHT.getCode());
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderBottom(BorderStyle.THIN);
        style.setLeftBorderColor(HSSFColorPredefined.BLACK.getIndex());
        style.setRightBorderColor(HSSFColorPredefined.BLACK.getIndex());
        style.setTopBorderColor(HSSFColorPredefined.BLACK.getIndex());
        style.setBottomBorderColor(HSSFColorPredefined.BLACK.getIndex());

        final CellStyle newStyle = workbook.createCellStyle();
        newStyle.cloneStyleFrom(style);
        this.setRowHeaderStyle(newStyle);
    }
}
