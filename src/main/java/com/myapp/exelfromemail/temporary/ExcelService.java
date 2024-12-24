package com.myapp.exelfromemail.temporary;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;

import java.io.InputStream;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

@Service
public class ExcelService {

//    public List<String[]> readExcel(InputStream inputStream) throws Exception {
//        List<String[]> data = new ArrayList<>();
//        Workbook workbook = new XSSFWorkbook(inputStream);
//        Sheet sheet = workbook.getSheetAt(0);
//
//        for (Row row : sheet) {
//            List<String> rowData = new ArrayList<>();
//            for (Cell cell : row) {
//                rowData.add(cell.toString());
//            }
//            data.add(rowData.toArray(new String[0]));
//        }
//        workbook.close();
//        return data;
//    }

    public Map<String, List<String>> readExcel(InputStream inputStream) throws Exception {
        Map<String, List<String>> data = new LinkedHashMap<>();

        try (Workbook workbook = new XSSFWorkbook(inputStream)) {
            Sheet sheet = workbook.getSheetAt(0); // Читаем первый лист

            // Получаем первую строку как заголовки
            Row headerRow = sheet.getRow(0);
            if (headerRow == null) {
                throw new Exception("Файл Excel не содержит заголовков.");
            }

            // Инициализируем ключи в Map по заголовкам
            for (Cell cell : headerRow) {
                data.put(cell.getStringCellValue(), new ArrayList<>());
            }

            // Заполняем данные начиная со второй строки
            for (int i = 1; i <= sheet.getLastRowNum(); i++) {
                Row row = sheet.getRow(i);
                if (row == null) continue;

                int cellIndex = 0;
                for (Cell cell : row) {
                    String header = headerRow.getCell(cellIndex).getStringCellValue();
                    List<String> columnData = data.get(header);

                    columnData.add(getCellValueAsString(cell));
                    cellIndex++;
                }
            }
        }

        return data;
    }

    private String getCellValueAsString(Cell cell) {
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

