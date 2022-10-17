package com.tiangang.poi.excel.template;

import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;
import lombok.NoArgsConstructor;
import org.apache.poi.ss.usermodel.*;

import java.util.ArrayList;
import java.util.Collection;
import java.util.List;

public class ExcelUtil {
    public static List<Cell> findCellList(Sheet sheet, List<String> findCellValueList) {
        List<Cell> cellList = new ArrayList<>();
        for (int i = sheet.getFirstRowNum(); i < sheet.getLastRowNum(); i++) {
            Row row = sheet.getRow(i);
            List<Cell> rowCellList = findCellList(row, findCellValueList);
            if (!isEmpty(rowCellList)) {
                cellList.addAll(rowCellList);
            }
        }
        return cellList;
    }

    public static List<Cell> findCellList(Row row, List<String> findCellValueList) {
        List<Cell> cellList = new ArrayList<>();
        for (int j = row.getFirstCellNum(); j < row.getLastCellNum(); j++) {
            String stringCellValue = row.getCell(j).getStringCellValue();
            if (findCellValueList.contains(stringCellValue)) {
                cellList.add(row.getCell(j));
            }
        }
        return cellList;
    }

    public static void setCellFont(List<Cell> cellList, FontParam fontParam) {
        if (isEmpty(cellList) || fontParam == null) {
            return;
        }
        CellStyle cellStyle = null;
        for (Cell cell : cellList) {
            if (cellStyle == null) {
                Workbook workbook = cell.getSheet().getWorkbook();
                cellStyle = workbook.createCellStyle();
                // 从现有样式克隆style，只修改Font，其它style不变
                cellStyle.cloneStyleFrom(cell.getCellStyle());
                // 获取原有字体
                Font oldFont = workbook.getFontAt(cellStyle.getFontIndexAsInt());
                // 创建新字体
                Font newFont = workbook.createFont();
                newFont.setFontName(fontParam.getFontName() == null? oldFont.getFontName(): fontParam.getFontName());
                newFont.setFontHeightInPoints(fontParam.getFontHeightInPoints() == null? oldFont.getFontHeightInPoints(): fontParam.getFontHeightInPoints());
                newFont.setBold(fontParam.getBold() == null? oldFont.getBold(): fontParam.getBold());
                newFont.setItalic(fontParam.getItalic() == null? oldFont.getItalic(): fontParam.getItalic());
                newFont.setStrikeout(fontParam.getStrikeout() == null? oldFont.getStrikeout(): fontParam.getStrikeout());
                newFont.setUnderline(fontParam.getUnderline() == null? oldFont.getUnderline(): fontParam.getUnderline());
                newFont.setColor(fontParam.getColor() == null? oldFont.getColor(): fontParam.getColor());
                // 设置字体
                cellStyle.setFont(newFont);
            }
            // 设置样式
            cell.setCellStyle(cellStyle);
        }
    }

    public static void hiddenColumn(Sheet sheet,int hiddenColumn){
        sheet.setColumnHidden(hiddenColumn,true);
    }

    private static boolean isEmpty(Collection<?> collection) {
        return (collection == null || collection.isEmpty());
    }

    /**
     * 字体参数类，为null代表不设置
     */
    @Data
    @Builder
    @NoArgsConstructor
    @AllArgsConstructor
    public static class FontParam {

        /**
         * 字体名
         */
        private String fontName;
        /**
         * 字体像素高度
         */
        private Short fontHeightInPoints;
        /**
         * 是否加粗
         */
        private Boolean bold;
        /**
         * 是否斜体
         */
        private Boolean italic;
        /**
         * 是否删除线
         */
        private Boolean strikeout;
        /**
         * 下划线类型
         * @see #U_NONE
         * @see #U_SINGLE
         * @see #U_DOUBLE
         * @see #U_SINGLE_ACCOUNTING
         * @see #U_DOUBLE_ACCOUNTING
         */
        private Byte underline;
        /**
         * 字体颜色
         */
        private Short color;

        /**
         * not underlined
         */
        public final static byte U_NONE = 0;

        /**
         * single (normal) underline
         */
        public final static byte U_SINGLE = 1;

        /**
         * double underlined
         */
        public final static byte U_DOUBLE = 2;

        /**
         * accounting style single underline
         */
        public final static byte U_SINGLE_ACCOUNTING = 0x21;

        /**
         * accounting style double underline
         */
        public final static byte U_DOUBLE_ACCOUNTING = 0x22;
    }
}
