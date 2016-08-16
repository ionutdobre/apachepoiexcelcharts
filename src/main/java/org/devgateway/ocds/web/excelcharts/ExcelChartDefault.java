package org.devgateway.ocds.web.excelcharts;

import org.apache.poi.ss.usermodel.Chart;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.charts.AxisCrosses;
import org.apache.poi.ss.usermodel.charts.AxisPosition;
import org.apache.poi.ss.usermodel.charts.ChartAxis;
import org.apache.poi.ss.usermodel.charts.ChartDataSource;
import org.apache.poi.ss.usermodel.charts.ChartLegend;
import org.apache.poi.ss.usermodel.charts.DataSources;
import org.apache.poi.ss.usermodel.charts.LegendPosition;
import org.apache.poi.ss.usermodel.charts.ValueAxis;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.devgateway.ocds.web.excelcharts.util.CustomChartDataFactory;
import org.devgateway.ocds.web.excelcharts.util.CustomChartDataFactoryDefault;

import java.util.List;

/**
 * @author idobre
 * @since 8/16/16
 */
public class ExcelChartDefault implements ExcelChart {
    private final ChartType type;

    private final List<String> categories;

    private final List<List<? extends Number>> values;

    private final Workbook workbook;

    public ExcelChartDefault(final ChartType type,
                             final List<String> categories,
                             final List<List<? extends Number>> values) {
        for (List<? extends Number> value : values) {
            if (categories.size() != value.size()) {
                throw new IllegalArgumentException("categories and value size should match!");
            }
        }

        this.type = type;
        this.categories = categories;
        this.values = values;
        this.workbook = new XSSFWorkbook();
    }

    @Override
    public Workbook createWorkbook() {
        final ExcelChartSheet excelChartSheet = new ExcelChartSheetDefault(workbook, type.toString());
        final Chart chart = excelChartSheet.createChartAndLegend();

        addCategories(excelChartSheet);
        addValues(excelChartSheet);

        final CustomChartDataFactory customChartDataFactory = new CustomChartDataFactoryDefault();
        final CustomChartData data = customChartDataFactory.createLineChartData("chart title");

        // Use a category axis for the bottom axis.
        final ChartAxis bottomAxis = chart.getChartAxisFactory().createCategoryAxis(AxisPosition.BOTTOM);
        final ValueAxis leftAxis = chart.getChartAxisFactory().createValueAxis(AxisPosition.LEFT);
        leftAxis.setCrosses(AxisCrosses.AUTO_ZERO);

        final ChartDataSource<String> categoryDataSource = excelChartSheet.getCategoryChartDataSource();
        final List<ChartDataSource<Number>> valuesDataSource = excelChartSheet.getValuesChartDataSource();

        for (ChartDataSource<Number> valueDataSource : valuesDataSource) {
            data.addSeries(categoryDataSource, valueDataSource);
        }

        chart.plot(data, bottomAxis, leftAxis);

        return workbook;
    }

    private void addCategories(ExcelChartSheet excelChartSheet) {
        final Row row = excelChartSheet.createRow();
        int coll = 0;
        for(String category : categories) {
            excelChartSheet.writeCell(category, row, coll++);
        }
    }

    private void addValues(ExcelChartSheet excelChartSheet) {
        Row row;
        for(List<? extends Number> value : values) {
            row = excelChartSheet.createRow();
            int coll = 0;
            for (Number val : value) {
                excelChartSheet.writeCell(val, row, coll++);
            }
        }
    }
}
