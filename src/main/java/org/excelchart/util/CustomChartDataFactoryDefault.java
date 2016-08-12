package org.excelchart.util;

import org.excelchart.areachart.AreaChartData;
import org.excelchart.areachart.XSSFAreaChartData;
import org.excelchart.barchart.BarChartData;
import org.excelchart.barchart.XSSFBarChartData;
import org.excelchart.linechart.LineChartData;
import org.excelchart.piechart.PieChartData;
import org.excelchart.piechart.XSSFPieChartData;
import org.excelchart.scatterchart.ScatterChartData;
import org.excelchart.scatterchart.XSSFScatterChartData;
import org.excelchart.stackedbarchart.StackedBarChartData;
import org.excelchart.stackedbarchart.XSSFStackedBarChartData;

/**
 * @author idobre
 * @since 8/11/16
 */
public class CustomChartDataFactoryDefault implements CustomChartDataFactory {
    /**
     * @return new scatter charts data instance
     */
    public ScatterChartData createScatterChartData() {
        return new XSSFScatterChartData();
    }

    /**
     * @return new line charts data instance
     */
    public LineChartData createLineChartData() {
        return new org.excelchart.linechart.XSSFLineChartData();
    }

    /**
     * @return new area charts data instance
     */
    public AreaChartData createAreaChartData() {
        return new XSSFAreaChartData();
    }

    /**
     * @return new bar charts data instance
     */
    public BarChartData createBarChartData() {
        return new XSSFBarChartData();
    }

    /**
     * @return new pie charts data instance
     */
    public PieChartData createPieChartData() {
        return new XSSFPieChartData();
    }

    /**
     * @return new stacked bar charts data instance
     */
    public StackedBarChartData createStackedBarChartData() {
        return new XSSFStackedBarChartData();
    }
}
