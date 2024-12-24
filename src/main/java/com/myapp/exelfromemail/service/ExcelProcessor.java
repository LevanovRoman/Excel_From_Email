package com.myapp.exelfromemail.service;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.*;

@Service
public class ExcelProcessor {

    public Map<String, List<String>> processExcelFile(File file) throws IOException {
        Map<String, List<String>> result = new HashMap<>();

        try (FileInputStream fis = new FileInputStream(file);
             Workbook workbook = new XSSFWorkbook(fis)) {

            Sheet sheet = workbook.getSheetAt(0); // Читаем первую страницу
            Row headerRow = sheet.getRow(0); // Заголовок (первая строка)

            if (headerRow == null) {
                throw new IllegalArgumentException("Файл не содержит заголовков");
            }

            // Инициализация Map ключей (заголовков)
            for (Cell cell : headerRow) {
                result.put(cell.getStringCellValue(), new ArrayList<>());
            }

            // Чтение данных
            for (int rowIndex = 1; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
                Row row = sheet.getRow(rowIndex);
                if (row == null) continue; // Пропускаем пустые строки

                int cellIndex = 0;
                for (Cell cell : row) {
                    String header = headerRow.getCell(cellIndex).getStringCellValue();
                    String cellValue = getCellValueAsString(cell);
                    result.get(header).add(cellValue);
                    cellIndex++;
                }
            }
        }

        return result;
    }

    private static String getCellValueAsString(Cell cell) {
        if (cell == null) return "";

        return switch (cell.getCellType()) {
            case STRING -> cell.getStringCellValue();
            case NUMERIC -> String.valueOf(cell.getNumericCellValue());
            case BOOLEAN -> String.valueOf(cell.getBooleanCellValue());
            case FORMULA -> cell.getCellFormula();
            default -> "";
        };
    }
}

