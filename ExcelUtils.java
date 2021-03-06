package com.optum.micro.intake;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import io.micrometer.core.instrument.util.TimeUtils;
import lombok.SneakyThrows;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.time.Duration;
import java.time.Instant;
import java.util.Map;
import java.util.Properties;
import java.util.concurrent.CompletableFuture;

public class ExcelUtils {


    @SneakyThrows
       public static void main(String[] args) {
        Workbook workbook = new XSSFWorkbook();

        Map<String, String> map = Map.of("Name", "sab", "Age", "30");
        Map<String, String> map1 = Map.of("Name", "sab1", "Age", "31");
        List<Map<String, String>> listMap = List.of(map, map1);

        Sheet sheet = workbook.createSheet("Persons");
        sheet.setColumnWidth(0, 6000);
        sheet.setColumnWidth(1, 4000);
        Row header = sheet.createRow(0);
        CellStyle headerStyle = workbook.createCellStyle();
        headerStyle.setFillForegroundColor(IndexedColors.LIGHT_BLUE.getIndex());
        headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        XSSFFont font = ((XSSFWorkbook) workbook).createFont();
        font.setFontName("Arial");
        font.setFontHeightInPoints((short) 16);
        font.setBold(true);
        headerStyle.setFont(font);

        Cell headerCell = header.createCell(0);
        headerCell.setCellValue("Name");
        headerCell.setCellStyle(headerStyle);

        headerCell = header.createCell(1);
        headerCell.setCellValue("Age");
        headerCell.setCellStyle(headerStyle);


        AtomicInteger rowNum = new AtomicInteger(1);
        
        
        listMap.forEach(tempMap -> {
            Row row = sheet.createRow(rowNum.get()); // increment the row here
            sheet.getRow(0).cellIterator().forEachRemaining(cell -> {
                if (tempMap.containsKey(cell.getStringCellValue())) {
                    Cell cell1 = row.createCell(cell.getColumnIndex());
                    cell1.setCellValue(tempMap.get(cell.getStringCellValue()));
                }
            });
            rowNum.getAndIncrement();
        });
        
        


        File currDir = new File(".");
        String path = currDir.getAbsolutePath();
        String fileLocation = path.substring(0, path.length() - 1) + "temp.xlsx";

        FileOutputStream outputStream = new FileOutputStream(fileLocation);
        workbook.write(outputStream);
        workbook.close();

    }
}
