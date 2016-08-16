package org.devgateway.ocds.web.excelcharts.util;

import org.devgateway.ocds.web.excelcharts.CustomChartData;

/**
 * @author idobre
 * @since 8/11/16
 *
 * A factory for different charts data types (like bar chart, area chart)
 */
public interface CustomChartDataFactory {
    /**
     * @return an appropriate ScatterChartData instance
     */
    CustomChartData createScatterChartData(final String title);

    /**
     * @return a LineChartData instance
     */
    CustomChartData createLineChartData(final String title);

    /**
     * @return a AreaChartData instance
     */
    CustomChartData createAreaChartData(final String title);

    /**
     * @return a BarChartData instance
     */
    CustomChartData createBarChartData(final String title);

    /**
     * @return a PieChartData instance
     */
    CustomChartData createPieChartData(final String title);

    /**
     * @return a StackedBarChartData instance
     */
    CustomChartData createStackedBarChartData(final String title);
}
