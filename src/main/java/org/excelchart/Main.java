package org.excelchart;

import org.apache.poi.ss.usermodel.Cell;
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
import org.apache.poi.ss.usermodel.charts.LineChartData;
import org.apache.poi.ss.usermodel.charts.ScatterChartData;
import org.apache.poi.ss.usermodel.charts.ValueAxis;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.excelchart.areachart.AreaChartData;
import org.excelchart.areachart.XSSFAreaChartData;
import org.excelchart.barchart.BarChartData;
import org.excelchart.barchart.XSSFBarChartData;
import org.excelchart.piechart.PieChartData;
import org.excelchart.piechart.XSSFPieChartData;
import stackedbarchart.StackedBarChartData;
import stackedbarchart.XSSFStackedBarChartData;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Random;

/**
 * @author idobre
 * @since 8/8/16
 */
public class Main {
    private static final int NUM_OF_ROWS = 3;
    private static final int NUM_OF_COLUMNS = 15;

    public static void main(String[] args) {
        System.out.println(">>> start");

        try {
            createScatterChart();
            createLineChart();
            createBarChart();
            createStackedBarChart();
            createAreaChartChart();
            createPieChartChart();
        } catch (IOException e) {
            e.printStackTrace();
        }
        System.out.println(">>> end");
    }

    private static Sheet createSheet(Workbook wb, String sheetName) {
        Sheet sheet = wb.createSheet(sheetName);

        // Create a row and put some cells in it. Rows are 0 based.
        Row row;
        Cell cell;
        for (int rowIndex = 0; rowIndex < NUM_OF_ROWS; rowIndex++) {
            row = sheet.createRow((short) rowIndex);
            for (int colIndex = 0; colIndex < NUM_OF_COLUMNS; colIndex++) {
                cell = row.createCell((short) colIndex);
                if (rowIndex == 0) {
                    cell.setCellValue(colIndex + 1);
                } else {
                    cell.setCellValue(getRandomNumberInRange(1, 10));
                }
            }
        }

        return sheet;
    }

    private static Chart createChartAndLegend(Sheet sheet) {
        Drawing drawing = sheet.createDrawingPatriarch();
        ClientAnchor anchor = drawing.createAnchor(0, 0, 0, 0, 0, 5, 15, 20);

        Chart chart = drawing.createChart(anchor);
        ChartLegend legend = chart.getOrCreateLegend();
        legend.setPosition(LegendPosition.TOP_RIGHT);

        return chart;
    }

    private static void createScatterChart() throws IOException {
        Workbook wb = new XSSFWorkbook();
        Sheet sheet = createSheet(wb, "scatter");

        Chart chart = createChartAndLegend(sheet);
        ScatterChartData data = chart.getChartDataFactory().createScatterChartData();

        ValueAxis bottomAxis = chart.getChartAxisFactory().createValueAxis(AxisPosition.BOTTOM);
        ValueAxis leftAxis = chart.getChartAxisFactory().createValueAxis(AxisPosition.LEFT);
        leftAxis.setCrosses(AxisCrosses.AUTO_ZERO);

        ChartDataSource<Number> xs = DataSources.fromNumericCellRange(sheet, new CellRangeAddress(0, 0, 0, NUM_OF_COLUMNS - 1));
        ChartDataSource<Number> ys1 = DataSources.fromNumericCellRange(sheet, new CellRangeAddress(1, 1, 0, NUM_OF_COLUMNS - 1));
        ChartDataSource<Number> ys2 = DataSources.fromNumericCellRange(sheet, new CellRangeAddress(2, 2, 0, NUM_OF_COLUMNS - 1));

        data.addSerie(xs, ys1);
        data.addSerie(xs, ys2);

        chart.plot(data, bottomAxis, leftAxis);

        // Write the output to a file
        FileOutputStream fileOut = new FileOutputStream("/Users/ionut/Downloads/ooxml-scatter-chart.xlsx");
        wb.write(fileOut);
        fileOut.close();
    }

    private static void createLineChart() throws IOException {
        Workbook wb = new XSSFWorkbook();
        Sheet sheet = createSheet(wb, "linechart");

        Chart chart = createChartAndLegend(sheet);
        LineChartData data = chart.getChartDataFactory().createLineChartData();

        // Use a category axis for the bottom axis.
        ChartAxis bottomAxis = chart.getChartAxisFactory().createCategoryAxis(AxisPosition.BOTTOM);
        ValueAxis leftAxis = chart.getChartAxisFactory().createValueAxis(AxisPosition.LEFT);
        leftAxis.setCrosses(AxisCrosses.AUTO_ZERO);

        ChartDataSource<Number> xs = DataSources.fromNumericCellRange(sheet, new CellRangeAddress(0, 0, 0, NUM_OF_COLUMNS - 1));
        ChartDataSource<Number> ys1 = DataSources.fromNumericCellRange(sheet, new CellRangeAddress(1, 1, 0, NUM_OF_COLUMNS - 1));
        ChartDataSource<Number> ys2 = DataSources.fromNumericCellRange(sheet, new CellRangeAddress(2, 2, 0, NUM_OF_COLUMNS - 1));

        data.addSeries(xs, ys1);
        data.addSeries(xs, ys2);

        chart.plot(data, bottomAxis, leftAxis);

        // Write the output to a file
        FileOutputStream fileOut = new FileOutputStream("/Users/ionut/Downloads/ooxml-line-chart.xlsx");
        wb.write(fileOut);
        fileOut.close();
    }

    private static void createBarChart() throws IOException {
        Workbook wb = new XSSFWorkbook();
        Sheet sheet = createSheet(wb, "barchart");

        Chart chart = createChartAndLegend(sheet);
        BarChartData data = new XSSFBarChartData();

        // Use a category axis for the bottom axis.
        ChartAxis bottomAxis = chart.getChartAxisFactory().createCategoryAxis(AxisPosition.BOTTOM);
        ValueAxis leftAxis = chart.getChartAxisFactory().createValueAxis(AxisPosition.LEFT);
        leftAxis.setCrosses(AxisCrosses.AUTO_ZERO);

        ChartDataSource<String> xs = DataSources.fromStringCellRange(sheet, new CellRangeAddress(0, 0, 0, NUM_OF_COLUMNS - 1));
        ChartDataSource<Number> ys1 = DataSources.fromNumericCellRange(sheet, new CellRangeAddress(1, 1, 0, NUM_OF_COLUMNS - 1));
        ChartDataSource<Number> ys2 = DataSources.fromNumericCellRange(sheet, new CellRangeAddress(2, 2, 0, NUM_OF_COLUMNS - 1));

        data.addSeries(xs, ys1);
        data.addSeries(xs, ys2);

        chart.plot(data, bottomAxis, leftAxis);

        // Write the output to a file
        FileOutputStream fileOut = new FileOutputStream("/Users/ionut/Downloads/ooxml-bar-chart.xlsx");
        wb.write(fileOut);
        fileOut.close();
    }

    private static void createAreaChartChart() throws IOException {
        Workbook wb = new XSSFWorkbook();
        Sheet sheet = createSheet(wb, "areachart");

        Chart chart = createChartAndLegend(sheet);
        AreaChartData data = new XSSFAreaChartData();

        // Use a category axis for the bottom axis.
        ChartAxis bottomAxis = chart.getChartAxisFactory().createCategoryAxis(AxisPosition.BOTTOM);
        ValueAxis leftAxis = chart.getChartAxisFactory().createValueAxis(AxisPosition.LEFT);
        leftAxis.setCrosses(AxisCrosses.AUTO_ZERO);

        ChartDataSource<String> xs = DataSources.fromStringCellRange(sheet, new CellRangeAddress(0, 0, 0, NUM_OF_COLUMNS - 1));
        ChartDataSource<Number> ys1 = DataSources.fromNumericCellRange(sheet, new CellRangeAddress(1, 1, 0, NUM_OF_COLUMNS - 1));

        data.addSeries(xs, ys1);

        chart.plot(data, bottomAxis, leftAxis);

        // Write the output to a file
        FileOutputStream fileOut = new FileOutputStream("/Users/ionut/Downloads/ooxml-area-chart.xlsx");
        wb.write(fileOut);
        fileOut.close();
    }

    private static void createPieChartChart() throws IOException {
        Workbook wb = new XSSFWorkbook();
        Sheet sheet = createSheet(wb, "piechart");

        Chart chart = createChartAndLegend(sheet);
        PieChartData data = new XSSFPieChartData();

        ChartDataSource<String> xs = DataSources.fromStringCellRange(sheet, new CellRangeAddress(0, 0, 0, NUM_OF_COLUMNS - 1));
        ChartDataSource<Number> ys1 = DataSources.fromNumericCellRange(sheet, new CellRangeAddress(1, 1, 0, NUM_OF_COLUMNS - 1));

        data.addSeries(xs, ys1);
        chart.plot(data);

        // Write the output to a file
        FileOutputStream fileOut = new FileOutputStream("/Users/ionut/Downloads/ooxml-pie-chart.xlsx");
        wb.write(fileOut);
        fileOut.close();
    }

    private static void createStackedBarChart() throws IOException {
        Workbook wb = new XSSFWorkbook();
        Sheet sheet = createSheet(wb, "stackedbarchart");

        Chart chart = createChartAndLegend(sheet);
        StackedBarChartData data = new XSSFStackedBarChartData();

        // Use a category axis for the bottom axis.
        ChartAxis bottomAxis = chart.getChartAxisFactory().createCategoryAxis(AxisPosition.BOTTOM);
        ValueAxis leftAxis = chart.getChartAxisFactory().createValueAxis(AxisPosition.LEFT);
        leftAxis.setCrosses(AxisCrosses.AUTO_ZERO);

        ChartDataSource<String> xs = DataSources.fromStringCellRange(sheet, new CellRangeAddress(0, 0, 0, NUM_OF_COLUMNS - 1));
        ChartDataSource<Number> ys1 = DataSources.fromNumericCellRange(sheet, new CellRangeAddress(1, 1, 0, NUM_OF_COLUMNS - 1));
        ChartDataSource<Number> ys2 = DataSources.fromNumericCellRange(sheet, new CellRangeAddress(2, 2, 0, NUM_OF_COLUMNS - 1));

        data.addSeries(xs, ys1);
        data.addSeries(xs, ys2);

        chart.plot(data, bottomAxis, leftAxis);

        // Write the output to a file
        FileOutputStream fileOut = new FileOutputStream("/Users/ionut/Downloads/ooxml-stackedbar-chart.xlsx");
        wb.write(fileOut);
        fileOut.close();
    }

    private static int getRandomNumberInRange(int min, int max) {
        Random r = new Random();
        return r.ints(min, (max + 1)).findFirst().getAsInt();
    }
}
