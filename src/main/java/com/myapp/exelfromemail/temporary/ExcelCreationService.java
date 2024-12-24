package com.myapp.exelfromemail.temporary;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;

import java.io.FileOutputStream;
import java.io.OutputStream;
import java.util.List;

@Service
public class ExcelCreationService {

    public void createExcel(List<String[]> data, String outputFilePath) throws Exception {
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Processed Data");

        for (int i = 0; i < data.size(); i++) {
            Row row = sheet.createRow(i);
            String[] rowData = data.get(i);
            for (int j = 0; j < rowData.length; j++) {
                Cell cell = row.createCell(j);
                cell.setCellValue(rowData[j]);
            }
        }

        try (OutputStream outputStream = new FileOutputStream(outputFilePath)) {
            workbook.write(outputStream);
        }
        workbook.close();
    }
}

