import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jfree.chart.ChartFactory;
import org.jfree.chart.ChartPanel;
import org.jfree.chart.ChartUtilities;
import org.jfree.chart.JFreeChart;
import org.jfree.data.time.*;

import javax.swing.*;
import java.awt.*;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.*;
import java.util.List;

import java.lang.Math.*;

import static java.lang.Math.pow;
import static java.lang.Math.sqrt;


public class Main {
    static Map<Integer, String> dates = new HashMap<Integer, String>();
    static Map<Integer, List<Double>> data = new HashMap<Integer, List<Double>>();
    static List<String> columnNames = new LinkedList<String>();
    static List<Double> averageValues = new LinkedList<Double>();
    static List<Double> dispersions = new LinkedList<Double>();
    static Map<Integer, List<Double>> correlations = new HashMap<Integer, List<Double>>();
    public static void main(String[] args) {

        try (XSSFWorkbook wb = new XSSFWorkbook("va_test_.xlsx")){
            //Reading xlsx
            XSSFSheet sheet = wb.getSheetAt(0);
            getColumnNames(sheet);
            for (int i = 1; i <= sheet.getLastRowNum(); i++) {
                getRowOfData(sheet, i);
            }
            //Making graphs
            createGraphWindow(createDataset());

            getZeroMean();
            System.out.println("zero mean done");
            getUnitVariance();
            System.out.println("unit variance done");
            getCorrelationMap();

            outResult("results.xlsx");

        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

    public static void getColumnNames(XSSFSheet sheet){
        XSSFRow row = sheet.getRow(0);
        for (int i = 0; i <= 16; i++){
            columnNames.add(row.getCell(i).getStringCellValue());
        }
    }
    public static void getRowOfData(XSSFSheet sheet, int index){
        XSSFRow row = sheet.getRow(index);
        List<Double> rowList = new LinkedList<Double>();
        if (row.getPhysicalNumberOfCells() != 17) {
            return;
        }
        dates.put(index, row.getCell(0).getStringCellValue());
        for (int i = 1; i <= 16; i++){
            rowList.add(Double.valueOf(row.getCell(i).getRawValue()));
        }
        data.put(index, rowList);
    }
    public static TimeSeries createSeries(int index){

        TimeSeries series = new TimeSeries(columnNames.get(index));
        SimpleDateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
        int count = 1;
        double value = 0;
        Date previous = new Date();
        for (Map.Entry<Integer, List<Double>> row : data.entrySet()) {
            String stringDate = dates.get(row.getKey());
            double currentValue = row.getValue().get(index);

            Date date = null;
            count++;
            try {
                date = dateFormat.parse(stringDate);
            } catch (ParseException e) {
                throw new RuntimeException(e);
            }

            value  += currentValue;
            if(date.getDay() != previous.getDay()){
                series.addOrUpdate(new Day( date), value/count);
                count = 1;
                value = 0;
            }
            previous = date;
        }
        return series;
    }
    public static TimeSeriesCollection createDataset(){
        TimeSeriesCollection dataset = new TimeSeriesCollection();
        for (int i = 0; i <= 14; i++){
            dataset.addSeries(createSeries(i));
        }
        return dataset;
    }
    public static void createGraphWindow(TimeSeriesCollection dataset) throws IOException {
        JFreeChart chart = ChartFactory.createTimeSeriesChart("Graphs", "", "",
                dataset, true, true, false);

        File outFile = new File("graph.jpeg");
        ChartUtilities.saveChartAsJPEG(outFile, chart, 800, 600);

        ChartPanel panel = new ChartPanel(chart);
        panel.setFillZoomRectangle(true);
        panel.setMouseZoomable(true);
        panel.setPreferredSize(new Dimension(800, 600));

        JFrame frame = new JFrame("Graph Example");
        frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        frame.getContentPane().add(panel);
        frame.pack();
        frame.setLocationRelativeTo(null);
        frame.setVisible(true);
    }
    public static void calculateAverageValues(){
        int count = 0;
        double value = 0;
        averageValues.add((double) 0);

        for (int i = 0; i <= 14; i++){
            for (List<Double> row : data.values()) {
                value += row.get(i);
                count++;

            }
            averageValues.add(i, value / count);
        }
    }
    public static void getZeroMean(){
        double value = 0;
        double averageValue = 0;
        calculateAverageValues();

        for (Map.Entry<Integer, List<Double>> row : data.entrySet()) {
            List<Double> newList = new LinkedList<Double>();
            for(int i = 0; i <= 14; i++){
                averageValue = averageValues.get(i);
                value = row.getValue().get(i);
                newList.add(value - averageValue);
            }
            newList.add(row.getValue().get(15));
            data.put(row.getKey(), newList);
        }
    }
    public static void calculateDispersion(){
        double value;
        double dispersion = 0;
        int count = 0;
        dispersions.add((double) 0);

        for (int i = 0; i <= 14; i++){
            for (List<Double> row : data.values()) {
                value = row.get(i);
                dispersion += value * value * count;
                count++;
            }
            dispersions.add(i, dispersion / count - pow(averageValues.get(i), 2));
        }
    }
    public static void getUnitVariance(){
        double value = 0;
        double dispersion = 0;
        calculateDispersion();

        for (Map.Entry<Integer, List<Double>> row : data.entrySet()) {
            List<Double> newList = new LinkedList<Double>();

            for(int i = 0; i <= 14; i++){
                dispersion = dispersions.get(i);
                value = row.getValue().get(i);
                newList.add(value / sqrt(dispersion));
            }
            newList.add(row.getValue().get(15));
            data.put(row.getKey(), newList);
        }
    }

    public static void getCorrelationMap(){

        for (int i = 0; i <= 14; i++){
            List<Double> correlationRow = new LinkedList<Double>();
            for (int j = 0; j <= 14; j++){
                if (i == j){
                    correlationRow.add(j, (double) 0);
                    continue;
                }
                correlationRow.add(j, getCorrelation(i, j));
            }
            correlations.put(i, correlationRow);
        }
    }
    public static double getCorrelation(int p1, int p2){
        double result = 0;

        result = (getXY(p1, p2) - averageValues.get(p1) * averageValues.get(p2)) / (sqrt(dispersions.get(p1)) * sqrt(dispersions.get(p2)));

        return result;
    }
    public static double getXY(int p1, int p2){
        double result = 0;
        int n = data.values().size();
        for (List<Double> row : data.values()){
            result += row.get(p1) * row.get(p2);
        }
        return result / n;
    }

    public static void outResult(String filename) throws IOException {
        XSSFWorkbook newWB = new XSSFWorkbook();
        XSSFSheet sheet = newWB.createSheet();

        XSSFRow firstRow = sheet.createRow(0);
        for (int i = 0; i <= 16; i++){
            XSSFCell cell = firstRow.createCell(i);
            cell.setCellValue(columnNames.get(i));
        }

        for (Map.Entry<Integer, List<Double>> row : data.entrySet()){
            XSSFRow newRow = sheet.createRow(row.getKey());
            XSSFCell dateCell = newRow.createCell(0);
            dateCell.setCellValue(dates.get(row.getKey()));
            for (int i = 0; i <= 15; i++){
                XSSFCell cell = newRow.createCell(i+1);
                cell.setCellValue(row.getValue().get(i));

            }
        }
        XSSFSheet sheet2 = newWB.createSheet("Correlation");
        XSSFRow fRow = sheet2.createRow(0);
        for (int i = 0; i <= 14; i++){
            XSSFRow row = sheet2.createRow(i+1);
            XSSFCell nameCell = fRow.createCell(i+1);
            XSSFCell nameCell2 = row.createCell(0);
            nameCell.setCellValue(columnNames.get(i+1));
            nameCell2.setCellValue(columnNames.get(i+1));
            for (int j = 0; j <= 14; j++){
                XSSFCell cell = row.createCell(j+1);
                cell.setCellValue(correlations.get(i).get(j));
            }
        }


        try (FileOutputStream fileout = new FileOutputStream(filename)) {
            newWB.write(fileout);
            System.out.println("Out ok");
        } catch (FileNotFoundException e) {
            throw new RuntimeException(e);
        }
    }

}
