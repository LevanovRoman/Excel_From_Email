package com.myapp.exelfromemail.controller;

import com.myapp.exelfromemail.service.ExcelProcessor;
import com.myapp.exelfromemail.service.ExcelReaderService;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import java.io.File;
import java.util.List;
import java.util.Map;

@RestController
@RequestMapping("api/get-excel")
public class HandlingExcelController {

    private final ExcelReaderService excelReaderService;
    private final ExcelProcessor excelProcessor;

    public HandlingExcelController(ExcelReaderService excelReaderService, ExcelProcessor excelProcessor) {
        this.excelReaderService = excelReaderService;
        this.excelProcessor = excelProcessor;
    }

    @GetMapping("/read-excel")
    public Map<String, List<String>> readExcelFromEmail() throws Exception {
        // Загрузка Excel из почты
        File excelFile = excelReaderService.downloadExcel("INBOX", "testexel");

        Map<String, List<String>> data = excelProcessor.processExcelFile(excelFile);
        // Вывод данных
        data.forEach((key, value) -> {
            System.out.println(key + ": " + value);
        });
        return data;
    }
}
