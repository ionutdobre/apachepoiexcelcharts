package org.devgateway.ocds.web.excelcharts.util;

import org.devgateway.ocds.web.excelcharts.CustomChartData;
import org.devgateway.ocds.web.excelcharts.data.XSSFAreaChartData;
import org.devgateway.ocds.web.excelcharts.data.XSSFBarChartData;
import org.devgateway.ocds.web.excelcharts.data.XSSFLineChartData;
import org.devgateway.ocds.web.excelcharts.data.XSSFPieChartData;
import org.devgateway.ocds.web.excelcharts.data.XSSFScatterChartData;
import org.devgateway.ocds.web.excelcharts.data.XSSFStackedBarChartData;

/**
 * @author idobre
 * @since 8/11/16
 */
public class CustomChartDataFactoryDefault implements CustomChartDataFactory {
    /**
     * @return new scatter charts data instance
     */
    @Override
    public CustomChartData createScatterChartData(final String title) {
        return new XSSFScatterChartData(title);
    }

    /**
     * @return new line charts data instance
     */
    @Override
    public CustomChartData createLineChartData(final String title) {
        return new XSSFLineChartData(title);
    }

    /**
     * @return new area charts data instance
     */
    @Override
    public CustomChartData createAreaChartData(final String title) {
        return new XSSFAreaChartData(title);
    }

    /**
     * @return new bar charts data instance
     */
    @Override
    public CustomChartData createBarChartData(final String title) {
        return new XSSFBarChartData(title);
    }

    /**
     * @return new pie charts data instance
     */
    @Override
    public CustomChartData createPieChartData(final String title) {
        return new XSSFPieChartData(title);
    }

    /**
     * @return new stacked bar charts data instance
     */
    @Override
    public CustomChartData createStackedBarChartData(final String title) {
        return new XSSFStackedBarChartData(title);
    }
}
