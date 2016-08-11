package org.excelchart.util;

import org.excelchart.areachart.AreaChartData;
import org.excelchart.barchart.BarChartData;
import org.excelchart.piechart.PieChartData;
import org.excelchart.stackedbarchart.StackedBarChartData;

/**
 * @author idobre
 * @since 8/11/16
 *
 * A factory for different charts data types (like bar chart, area chart)
 */
public interface CustomChartDataFactory {
    /**
     * @return a AreaChartData instance
     */
    AreaChartData createAreaChartData();

    /**
     * @return a BarChartData instance
     */
    BarChartData createBarChartData();

    /**
     * @return a PieChartData instance
     */
    PieChartData createPieChartData();

    /**
     * @return a StackedBarChartData instance
     */
    StackedBarChartData createStackedBarChartData();
}
