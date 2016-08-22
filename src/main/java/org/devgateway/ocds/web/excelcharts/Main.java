package org.devgateway.ocds.web.excelcharts;

import org.apache.poi.ss.usermodel.Workbook;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Arrays;
import java.util.List;

/**
 * @author idobre
 * @since 8/8/16
 */
public class Main {
    private static final List<?> categories = Arrays.asList(
            "cat 1",
            "cat 2",
            "cat 3",
            "cat 4",
            "cat 5"
    );

    private static final List<List<? extends Number>> values = Arrays.asList(
            Arrays.asList(5, 7, 10, 12, 6),
            Arrays.asList(20, 12, 10, 5, 14)
    );

    public static void main(final String[] args) {
        System.out.println(">>> start");

        try {
            createScatterChart();
            createLineChart();
            createBarChart();
            createStackedBarChart();
            createAreaChartChart();
            createPieChartChart();
            createBubblehartChart();
        } catch (IOException e) {
            e.printStackTrace();
        }

        System.out.println(">>> end");
    }

    private static void createScatterChart() throws IOException {
        final ExcelChart excelChart = new ExcelChartDefault("scatter chart", ChartType.scatter, categories, values);
        final Workbook workbook = excelChart.createWorkbook();

        // Write the output to a file
        FileOutputStream fileOut = new FileOutputStream("/Users/ionut/Downloads/ooxml-scatter-chart.xlsx");
        workbook.write(fileOut);
        fileOut.close();
    }

    private static void createLineChart() throws IOException {
        final ExcelChart excelChart = new ExcelChartDefault("line chart", ChartType.line, categories, values);
        excelChart.configureSeriesTitle(
                Arrays.asList(
                        "foo",
                        "bar"
                ));
        final Workbook workbook = excelChart.createWorkbook();

        // Write the output to a file
        FileOutputStream fileOut = new FileOutputStream("/Users/ionut/Downloads/ooxml-line-chart.xlsx");
        workbook.write(fileOut);
        fileOut.close();
    }

    private static void createBarChart() throws IOException {
        final ExcelChart excelChart = new ExcelChartDefault("bar chart", ChartType.bar, categories, values);
        final Workbook workbook = excelChart.createWorkbook();

        // Write the output to a file
        FileOutputStream fileOut = new FileOutputStream("/Users/ionut/Downloads/ooxml-bar-chart.xlsx");
        workbook.write(fileOut);
        fileOut.close();
    }

    private static void createAreaChartChart() throws IOException {
        final ExcelChart excelChart = new ExcelChartDefault("area chart", ChartType.area, categories,
                values.subList(0, 1));
        final Workbook workbook = excelChart.createWorkbook();

        // Write the output to a file
        FileOutputStream fileOut = new FileOutputStream("/Users/ionut/Downloads/ooxml-area-chart.xlsx");
        workbook.write(fileOut);
        fileOut.close();
    }

    private static void createPieChartChart() throws IOException {
        final ExcelChart excelChart = new ExcelChartDefault("pie chart", ChartType.pie,
                categories, values.subList(0, 1));
        final Workbook workbook = excelChart.createWorkbook();

        // Write the output to a file
        FileOutputStream fileOut = new FileOutputStream("/Users/ionut/Downloads/ooxml-pie-chart.xlsx");
        workbook.write(fileOut);
        fileOut.close();
    }

    private static void createStackedBarChart() throws IOException {
        final ExcelChart excelChart = new ExcelChartDefault("stacked chart", ChartType.stacked, categories, values);
        final Workbook workbook = excelChart.createWorkbook();

        // Write the output to a file
        FileOutputStream fileOut = new FileOutputStream("/Users/ionut/Downloads/ooxml-stackedbar-chart.xlsx");
        workbook.write(fileOut);
        fileOut.close();
    }

    private static void createBubblehartChart() throws IOException {
        final ExcelChart excelChart = new ExcelChartDefault("bubble chart", ChartType.bubble,
                categories, values.subList(0, 1));
        final Workbook workbook = excelChart.createWorkbook();

        // Write the output to a file
        FileOutputStream fileOut = new FileOutputStream("/Users/ionut/Downloads/ooxml-bubble-chart.xlsx");
        workbook.write(fileOut);
        fileOut.close();
    }
}
