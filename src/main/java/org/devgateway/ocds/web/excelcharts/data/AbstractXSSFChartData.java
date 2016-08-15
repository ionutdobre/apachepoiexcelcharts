package org.devgateway.ocds.web.excelcharts.data;

import org.apache.poi.ss.usermodel.charts.ChartDataSource;
import org.devgateway.ocds.web.excelcharts.CustomChartData;
import org.devgateway.ocds.web.excelcharts.CustomChartSeries;

import java.util.ArrayList;
import java.util.List;

/**
 * @author idobre
 * @since 8/12/16
 */
public abstract class AbstractXSSFChartData implements CustomChartData {
    /**
     * List of all data series.
     */
    protected final List<CustomChartSeries> series;

    public AbstractXSSFChartData() {
        series = new ArrayList<>();
    }

    @Override
    public CustomChartSeries addSeries(final ChartDataSource<?> categoryAxisData,
                                       final ChartDataSource<? extends Number> values) {
        if (!values.isNumeric()) {
            throw new IllegalArgumentException("Value data source must be numeric.");
        }

        int numOfSeries = series.size();
        CustomChartSeries newSeries = createNewSerie(numOfSeries, numOfSeries, categoryAxisData, values);
        series.add(newSeries);

        return newSeries;
    }

    /**
     * Add a new Serie specific to each AbstractXSSFChartData implementation.
     */
    protected abstract CustomChartSeries createNewSerie(final int id, final int order,
                                                        final ChartDataSource<?> categories,
                                                        final ChartDataSource<? extends Number> values);
}
