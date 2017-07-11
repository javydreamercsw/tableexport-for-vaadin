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

/**
 * Example of how the ExcelExport class might be extended to implement specific
 * formatting features in the exported file.
 */
public class EnhancedFormatExcelExport extends ExcelExport {

    /**
     * The Constant serialVersionUID.
     */
    private static final long serialVersionUID = 9113961084041090666L;

    public EnhancedFormatExcelExport(final Table table) {
        this(table, "Enhanced Export");
    }

    public EnhancedFormatExcelExport(final TableHolder tableHolder) {
        this(tableHolder, "Enhanced Export");
    }

    public EnhancedFormatExcelExport(final TableHolder tableHolder,
            final String sheetName) {
        super(tableHolder, sheetName);
        format();
    }

    public EnhancedFormatExcelExport(final Table table, final String sheetName) {
        super(table, sheetName);
        format();
    }

    private void format() {
        this.setRowHeaders(true);
        CellStyle style;
        Font f;

        style = this.getTitleStyle();
        setStyle(style, HSSFColorPredefined.DARK_BLUE.getIndex(), 18,
                HSSFColorPredefined.WHITE.getIndex(),
                true,
                HorizontalAlignment.CENTER_SELECTION.getCode());

        style = this.getColumnHeaderStyle();
        setStyle(style, HSSFColorPredefined.LIGHT_BLUE.getIndex(), 12,
                HSSFColorPredefined.BLACK.getIndex(),
                true,
                HorizontalAlignment.CENTER_SELECTION.getCode());

        style = this.getDateDataStyle();
        setStyle(style, HSSFColorPredefined.LIGHT_CORNFLOWER_BLUE.getIndex(), 12,
                HSSFColorPredefined.BLACK.getIndex(), false,
                HorizontalAlignment.RIGHT.getCode());

        style = this.getDoubleDataStyle();
        setStyle(style, HSSFColorPredefined.LIGHT_CORNFLOWER_BLUE.getIndex(), 12,
                HSSFColorPredefined.BLACK.getIndex(), false,
                HorizontalAlignment.RIGHT.getCode());
        this.setTotalsDoubleStyle(style);

        style = this.getIntegerDataStyle();
        setStyle(style, HSSFColorPredefined.LIGHT_CORNFLOWER_BLUE.getIndex(), 12,
                HSSFColorPredefined.BLACK.getIndex(), false,
                HorizontalAlignment.RIGHT.getCode());
        this.setTotalsIntegerStyle(style);

        // we want the rowHeader style to be like the columnHeader style, just centered differently.
        final CellStyle newStyle = workbook.createCellStyle();
        newStyle.cloneStyleFrom(style);
        newStyle.setAlignment(HorizontalAlignment.LEFT.getCode());
        this.setRowHeaderStyle(newStyle);
    }

    private void setStyle(CellStyle style, short foregroundColor,
            int fontHeight, short fontColor,
            boolean fontBoldweight, short alignment) {
        style.setFillForegroundColor(foregroundColor);
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND.getCode());
        Font f = workbook.getFontAt(style.getFontIndex());
        f.setFontHeightInPoints((short) fontHeight);
        f.setFontName(HSSFFont.FONT_ARIAL);
        f.setColor(fontColor);
        f.setBold(fontBoldweight);
        style.setAlignment(alignment);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderBottom(BorderStyle.THIN);
        style.setLeftBorderColor(HSSFColorPredefined.BLACK.getIndex());
        style.setRightBorderColor(HSSFColorPredefined.BLACK.getIndex());
        style.setTopBorderColor(HSSFColorPredefined.BLACK.getIndex());
        style.setBottomBorderColor(HSSFColorPredefined.BLACK.getIndex());
    }

}
