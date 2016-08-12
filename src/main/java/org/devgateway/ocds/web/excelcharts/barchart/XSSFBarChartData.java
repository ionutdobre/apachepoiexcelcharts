package org.devgateway.ocds.web.excelcharts.barchart;

import org.apache.poi.ss.usermodel.Chart;
import org.apache.poi.ss.usermodel.charts.ChartAxis;
import org.apache.poi.ss.usermodel.charts.ChartDataSource;
import org.apache.poi.xssf.usermodel.XSSFChart;
import org.apache.poi.xssf.usermodel.charts.AbstractXSSFChartSeries;
import org.devgateway.ocds.web.excelcharts.util.XSSFChartUtil;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTAxDataSource;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTBarChart;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTBarSer;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTNumDataSource;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTPlotArea;
import org.openxmlformats.schemas.drawingml.x2006.chart.STBarDir;

import java.util.ArrayList;
import java.util.List;

/**
 * @author idobre
 * @since 8/8/16
 * Holds data for a XSSF Bar Chart
 */
public class XSSFBarChartData implements BarChartData {
    /**
     * List of all data series.
     */
    private List<XSSFBarChartData.Series> series;

    public XSSFBarChartData() {
        series = new ArrayList<XSSFBarChartData.Series>();
    }

    static class Series extends AbstractXSSFChartSeries implements BarChartSeries {
        private int id;
        private int order;
        private ChartDataSource<?> categories;
        private ChartDataSource<? extends Number> values;

        protected Series(int id, int order,
                         ChartDataSource<?> categories,
                         ChartDataSource<? extends Number> values) {
            this.id = id;
            this.order = order;
            this.categories = categories;
            this.values = values;
        }

        public ChartDataSource<?> getCategoryAxisData() {
            return categories;
        }

        public ChartDataSource<? extends Number> getValues() {
            return values;
        }

        protected void addToChart(CTBarChart ctBarChart) {
            CTBarSer ctBarSer = ctBarChart.addNewSer();
            ctBarSer.addNewIdx().setVal(id);
            ctBarSer.addNewOrder().setVal(order);

            CTAxDataSource catDS = ctBarSer.addNewCat();
            XSSFChartUtil.buildAxDataSource(catDS, categories);

            CTNumDataSource valueDS = ctBarSer.addNewVal();
            XSSFChartUtil.buildNumDataSource(valueDS, values);

            if (isTitleSet()) {
                ctBarSer.setTx(getCTSerTx());
            }
        }
    }

    public BarChartSeries addSeries(ChartDataSource<?> categoryAxisData, ChartDataSource<? extends Number> values) {
        if (!values.isNumeric()) {
            throw new IllegalArgumentException("Value data source must be numeric.");
        }
        int numOfSeries = series.size();
        XSSFBarChartData.Series newSeries =
                new XSSFBarChartData.Series(numOfSeries, numOfSeries, categoryAxisData, values);
        series.add(newSeries);
        return newSeries;
    }

    public List<? extends BarChartSeries> getSeries() {
        return series;
    }

    public void fillChart(Chart chart, ChartAxis... axis) {
        if (!(chart instanceof XSSFChart)) {
            throw new IllegalArgumentException("Chart must be instance of XSSFChart");
        }

        XSSFChart xssfChart = (XSSFChart) chart;
        CTPlotArea plotArea = xssfChart.getCTChart().getPlotArea();
        CTBarChart barChart = plotArea.addNewBarChart();

        barChart.addNewVaryColors().setVal(false);

        // set bars orientation
        barChart.addNewBarDir().setVal(STBarDir.COL);

        for (XSSFBarChartData.Series s : series) {
            s.addToChart(barChart);
        }

        for (ChartAxis ax : axis) {
            barChart.addNewAxId().setVal(ax.getId());
        }
    }
}