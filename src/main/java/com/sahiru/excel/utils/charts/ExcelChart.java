package com.sahiru.excel.utils.charts;

import com.sahiru.excel.utils.data.ChartData;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;
import java.util.Set;

public class ExcelChart {

    private XSSFWorkbook workbook;
    private XSSFSheet sheet;

    public ExcelChart(File file) throws NullPointerException, InvalidFormatException, IOException {
        if(file == null) {
            throw new NullPointerException("File cannot be null");
        }

        this.workbook = new XSSFWorkbook(OPCPackage.open(file));
        this.sheet = workbook.getSheetAt(0);
    }

    public ExcelChart drawChart(List<ChartData> dataList, int startRowIndex) {

        for (ChartData chartData : dataList) {
            Set<String> keys = chartData.keySet();
            int rowIndex = startRowIndex;
            for(String key : keys) {
                XSSFRow row = this.sheet.getRow(rowIndex);
                if(row == null) {
                    row = this.sheet.createRow(rowIndex);
                }
                row.createCell(chartData.getLabelColumnIndex()).setCellValue(key);
                row.createCell(chartData.getValueColumnIndex()).setCellValue(Integer.valueOf(chartData.get(key).toString()));
                ++rowIndex;
            }
        }
        return this;
    }

    public void writeToFile() throws IOException {
        FileOutputStream fileOut = new FileOutputStream("E:\\test.xlsx");
        this.workbook.write(fileOut);
    }
}
