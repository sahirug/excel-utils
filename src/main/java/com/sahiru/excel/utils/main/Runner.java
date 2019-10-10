package com.sahiru.excel.utils.main;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Date;

public class Runner {

    public void loadExcel() throws IOException, InvalidFormatException {
        File file = new File(getClass().getClassLoader().getResource("finalExcelMetricTemplate.xlsx").getFile());
        XSSFWorkbook workbook = new XSSFWorkbook(OPCPackage.open(file));
        XSSFSheet sheet = workbook.getSheetAt(0);

        XSSFRow row0 = sheet.createRow(0);
        XSSFRow row1 = sheet.createRow(8);
        XSSFRow row2 = sheet.createRow(9);
        XSSFRow row3 = sheet.createRow(10);

        row0.createCell(11).setCellValue("Company Name");
        row0.getCell(11).getCellStyle().setAlignment(HorizontalAlignment.CENTER);

        row1.createCell(0).setCellValue("AAAA");
        row1.createCell(1).setCellValue(24);
        row1.createCell(11).setCellValue("EEEE");
        row1.createCell(14).setCellValue(33);
        row1.createCell(28).setCellValue("YYYY");
        row1.createCell(31).setCellValue(21);

        row2.createCell(0).setCellValue("IOI");
        row2.createCell(1).setCellValue(12);
        row2.createCell(11).setCellValue("LPL");
        row2.createCell(14).setCellValue(45);
        row2.createCell(28).setCellValue("FTF");
        row2.createCell(31).setCellValue(76);

        row3.createCell(0).setCellValue("DSDS");
        row3.createCell(1).setCellValue(22);
        row3.createCell(11).setCellValue("SSSA");
        row3.createCell(14).setCellValue(9);
        row3.createCell(28).setCellValue("FEFE");
        row3.createCell(31).setCellValue(13);

        FileOutputStream fileOut = new FileOutputStream("E:\\test.xlsx");
        workbook.write(fileOut);

        System.out.println("done");
    }

    public static void main(String[] args) throws IOException, InvalidFormatException {
        Runner r = new Runner();
        r.loadExcel();
    }
}
