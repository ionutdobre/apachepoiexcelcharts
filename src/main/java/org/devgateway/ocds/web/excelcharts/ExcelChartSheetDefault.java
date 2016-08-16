package org.devgateway.ocds.web.excelcharts;

import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Chart;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.charts.ChartDataSource;
import org.apache.poi.ss.usermodel.charts.ChartLegend;
import org.apache.poi.ss.usermodel.charts.DataSources;
import org.apache.poi.ss.usermodel.charts.LegendPosition;
import org.apache.poi.ss.util.CellRangeAddress;

import java.math.BigDecimal;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Collection;
import java.util.Date;
import java.util.List;
import java.util.Set;

/**
 * Class that prepares the default Styles and Fonts for Excel cells.
 *
 * @author idobre
 * @since 8/16/16
 */
public final class ExcelChartSheetDefault implements ExcelChartSheet {
    private final Workbook workbook;

    private final Sheet excelSheet;

    private Font dataFont;

    private Font headerFont;

    private final CellStyle dataStyleCell;

    private final CellStyle headerStyleCell;

    // declare only one cell object reference
    private Cell cell;

    public ExcelChartSheetDefault(final Workbook workbook, final String excelSheetName) {
        this.workbook = workbook;
        this.excelSheet = workbook.createSheet(excelSheetName);
        // this.excelSheet.setColumnWidth(0, 5000); - TODO

        // get the styles from workbook without creating them again (by default the workbook has already 1 style)
        if (workbook.getNumCellStyles() > 1) {
            this.dataStyleCell = workbook.getCellStyleAt((short) 1);
            this.headerStyleCell = workbook.getCellStyleAt((short) 2);
        } else {
            // init the fonts and styles
            this.dataFont = this.workbook.createFont();
            this.dataFont.setFontHeightInPoints((short) 12);
            this.dataFont.setFontName("Times New Roman");
            this.dataFont.setColor(HSSFColor.BLACK.index);

            this.headerFont = this.workbook.createFont();
            this.headerFont.setFontHeightInPoints((short) 14);
            this.headerFont.setFontName("Times New Roman");
            this.headerFont.setColor(HSSFColor.BLACK.index);
            this.headerFont.setBold(true);

            this.dataStyleCell = this.workbook.createCellStyle();
            this.dataStyleCell.setAlignment(CellStyle.ALIGN_LEFT);
            this.dataStyleCell.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
            this.dataStyleCell.setWrapText(true);
            this.dataStyleCell.setFont(this.dataFont);

            this.headerStyleCell = this.workbook.createCellStyle();
            this.headerStyleCell.setAlignment(CellStyle.ALIGN_CENTER);
            this.headerStyleCell.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
            this.headerStyleCell.setWrapText(true);
            this.headerStyleCell.setFont(this.headerFont);
        }
    }

    /**
     * Creates a cell and tries to determine it's type based on the value type
     *
     * there is only one Cell object otherwise the Heap Space will fill really quickly
     *
     * @param value
     * @param row
     * @param column
     */
    @Override
    public void writeCell(final Object value, final Row row, final int column) {
        // try to determine the cell type based on the object value
        // if nothing matches then use 'CELL_TYPE_STRING' as type and call the object toString() function.
        //      * don't create any cell if the value is null (Cell.CELL_TYPE_BLANK)
        //      * do nothing if we have an empty List/Set instead of display empty brackets like []
        if (value != null && !((value instanceof List || value instanceof Set) && ((Collection) value).isEmpty())) {
            if (value instanceof String) {
                cell = row.createCell(column, Cell.CELL_TYPE_STRING);
                cell.setCellValue((String) value);
            } else {
                if (value instanceof Integer) {
                    cell = row.createCell(column, Cell.CELL_TYPE_NUMERIC);
                    cell.setCellValue((Integer) value);
                } else {
                    if (value instanceof BigDecimal) {
                        cell = row.createCell(column, Cell.CELL_TYPE_NUMERIC);
                        cell.setCellValue(((BigDecimal) value).doubleValue());
                    } else {
                        if (value instanceof Boolean) {
                            cell = row.createCell(column, Cell.CELL_TYPE_BOOLEAN);
                            cell.setCellValue(((Boolean) value) ? "Yes" : "No");
                        } else {
                            if (value instanceof Date) {
                                SimpleDateFormat sdf = new SimpleDateFormat("dd-MM-yyyy");

                                cell = row.createCell(column, Cell.CELL_TYPE_STRING);
                                cell.setCellValue(sdf.format((Date) value));
                            } else {
                                cell = row.createCell(column, Cell.CELL_TYPE_STRING);
                                cell.setCellValue(value.toString());
                            }
                        }
                    }
                }
            }

            // determine the style of the row based on it's index
            if (row.getRowNum() < 1) {
                cell.setCellStyle(headerStyleCell);
            } else {
                cell.setCellStyle(dataStyleCell);
            }
        } else {
            // create a Cell.CELL_TYPE_BLANK
            row.createCell(column);
        }
    }

    /**
     * Create a new row and set the default height (different heights for headers and data rows)
     *
     * @param rowNumber
     * @return Row
     */
    public Row createRow(final int rowNumber) {
        Row row = excelSheet.createRow(rowNumber);

        if (rowNumber < 1) {
            row.setHeight((short) 800);             // 40px (800 / 10 / 2)
        } else {
            row.setHeight((short) 600);             // 30px (600 / 10 / 2)
        }

        return row;
    }

    /**
     * Create a new row and return it. Since the rows in the sheet are 0-based we can use
     * {@link Sheet#getPhysicalNumberOfRows} to get the new free row
     */
    public Row createRow() {
        return createRow(excelSheet.getPhysicalNumberOfRows());
    }

    /**
     * Creates a chart and also attaches a legend to it.
     */
    public Chart createChartAndLegend() {
        final Drawing drawing = excelSheet.createDrawingPatriarch();
        final ClientAnchor anchor = drawing.createAnchor(0, 0, 0, 0, 0, 5, 15, 20);
        final Chart chart = drawing.createChart(anchor);

        final ChartLegend legend = chart.getOrCreateLegend();
        legend.setPosition(LegendPosition.BOTTOM);

        return chart;
    }

    @Override
    public ChartDataSource<String> getCategoryChartDataSource() {
        return getChartDataSource(0);      // categories should always be on the first row
    }

    @Override
    public List<ChartDataSource<Number>> getValuesChartDataSource() {
        final List<ChartDataSource<Number>> valuesDataSource = new ArrayList<>();

        // values should always be after the first row (rows are 0-based)
        for(int i = 1; i < excelSheet.getPhysicalNumberOfRows(); i++) {
            valuesDataSource.add(getChartDataSource(i));
        }

        return valuesDataSource;
    }

    private ChartDataSource getChartDataSource(int row) {
        final int lastCellNum = excelSheet.getRow(row).getLastCellNum() - 1;
        final CellRangeAddress cellRangeAddress = new CellRangeAddress(row, row, 0, lastCellNum);
        if (row == 0) {
            return DataSources.fromStringCellRange(excelSheet, cellRangeAddress);
        } else {
            return DataSources.fromNumericCellRange(excelSheet, cellRangeAddress);
        }
    }
}
