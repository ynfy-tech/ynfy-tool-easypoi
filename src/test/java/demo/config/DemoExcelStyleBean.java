package demo.config;

import cn.afterturn.easypoi.excel.entity.params.ExcelExportEntity;
import cn.afterturn.easypoi.excel.entity.params.ExcelForEachParams;
import cn.afterturn.easypoi.excel.export.styler.IExcelExportStyler;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.usermodel.Font;

/**
 * 用于定义 poi 样式
 * @author Hsiong
 * @version 1.0.0
 * @desc 〈〉
 */
public class DemoExcelStyleBean implements IExcelExportStyler {
    private static final short STRING_FORMAT = (short) BuiltinFormats.getBuiltinFormat("TEXT");
    private static final short FONT_SIZE_TEN = 9;
    private static final short FONT_SIZE_ELEVEN = 10;
    private static final short FONT_SIZE_TWELVE = 10;
    /**
     * 大标题样式
     */
    private CellStyle headerStyle;
    /**
     * 每列标题样式
     */
    private CellStyle titleStyle;
    /**
     * 数据行样式
     */
    private CellStyle styles;

    public DemoExcelStyleBean(Workbook workbook) {
        this.init(workbook);
    }

    /**
     * 初始化样式
     *
     * @param workbook
     */
    private void init(Workbook workbook) {
        this.headerStyle = initHeaderStyle(workbook);
        this.titleStyle = initTitleStyle(workbook);
        this.styles = initStyles(workbook);
    }

    /**
     * 大标题样式
     *
     * @param color
     * @return
     */
    @Override
    public CellStyle getHeaderStyle(short color) {
        return headerStyle;
    }

    /**
     * 每列标题样式
     *
     * @param color
     * @return
     */
    @Override
    public CellStyle getTitleStyle(short color) {
        return titleStyle;
    }

    /**
     * 数据行样式
     *
     * @param parity 可以用来表示奇偶行
     * @param entity 数据内容
     * @return 样式
     */
    @Override
    public CellStyle getStyles(boolean parity, ExcelExportEntity entity) {
        return styles;
    }

    @Override
    public CellStyle getStyles(Cell cell, int i, ExcelExportEntity excelExportEntity, Object o, Object o1) {
        return styles;
    }

    @Override
    public CellStyle getTemplateStyles(boolean b, ExcelForEachParams excelForEachParams) {
        return styles;
    }

    /**
     * 初始化--大标题样式
     *
     * @param workbook
     * @return
     */
    private CellStyle initHeaderStyle(Workbook workbook) {
        CellStyle style = getBaseCellStyle(workbook);
        style.setFont(getFont(workbook, FONT_SIZE_TWELVE, true, IndexedColors.BLUE.getIndex()));
        style.setFillForegroundColor(IndexedColors.SKY_BLUE.getIndex());
        style.setFillPattern(CellStyle.SOLID_FOREGROUND);
        return style;
    }

    /**
     * 初始化--每列标题样式
     *
     * @param workbook
     * @return
     */
    private CellStyle initTitleStyle(Workbook workbook) {
        CellStyle style = getBaseCellStyle(workbook);
        style.setFont(getFont(workbook, FONT_SIZE_ELEVEN, false, IndexedColors.BLUE.getIndex()));
        //背景色
        style.setFillForegroundColor(IndexedColors.PALE_BLUE.getIndex());
        style.setFillPattern(CellStyle.SOLID_FOREGROUND);
        return style;
    }

    /**
     * 初始化--数据行样式
     *
     * @param workbook
     * @return
     */
    private CellStyle initStyles(Workbook workbook) {
        CellStyle style = getBaseCellStyle(workbook);
        style.setFont(getFont(workbook, FONT_SIZE_TEN, false, IndexedColors.BLUE.getIndex()));
        style.setFillForegroundColor(IndexedColors.LIGHT_YELLOW.getIndex());
        style.setFillPattern(CellStyle.SOLID_FOREGROUND);
        //style.setDataFormat(STRING_FORMAT);
        return style;
    }

    /**
     * 基础样式
     *
     * @return
     */
    private CellStyle getBaseCellStyle(Workbook workbook) {
        CellStyle style = workbook.createCellStyle();
        //下边框
        style.setBorderBottom(CellStyle.BORDER_THIN);
        //左边框
        style.setBorderLeft(CellStyle.BORDER_THIN);
        //上边框
        style.setBorderTop(CellStyle.BORDER_MEDIUM);
        //右边框
        style.setBorderRight(CellStyle.BORDER_THIN);
        //水平居中
        style.setAlignment(CellStyle.ALIGN_CENTER);
        //上下居中
        style.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
        //设置自动换行
        style.setWrapText(true);
        return style;
    }

    /**
     * 字体样式
     *
     * @param size   字体大小
     * @param isBold 是否加粗
     * @return
     */
    private Font getFont(Workbook workbook, short size, boolean isBold, short color) {
        Font font = workbook.createFont();
        //字体样式
        font.setFontName("宋体");
        //是否加粗
        font.setBold(isBold);
        //字体大小
        font.setFontHeightInPoints(size);
        font.setColor(color);
        return font;
    }

}
