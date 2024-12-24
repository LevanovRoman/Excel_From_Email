package com.myapp.exelfromemail.temporary;

import com.myapp.exelfromemail.service.*;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RestController;

import java.io.InputStream;
import java.util.List;
import java.util.Map;

@RestController
public class FileController {

    private final EmailService emailService;
    private final ExcelService excelService;
    private final ExcelCreationService excelCreationService;
    private final ExcelReaderService excelReaderService;
    private final ExcelProcessor excelProcessor;

    public FileController(EmailService emailService, ExcelService excelService,
                          ExcelCreationService excelCreationService, ExcelReaderService excelReaderService, ExcelProcessor excelProcessor) {
        this.emailService = emailService;
        this.excelService = excelService;
        this.excelCreationService = excelCreationService;
        this.excelReaderService = excelReaderService;
        this.excelProcessor = excelProcessor;
    }

    @GetMapping("/process-email")
    public String processEmail() {
        try {
            // Получение Excel-файла из почты
            InputStream inputStream = emailService.getExcelAttachment(
                    "cgp.nordsy.spb.ru", "imap", "10390@nordsy.spb.ru", "rfnfvfhfy342A"
//                    "cgp.nordsy.spb.ru", "imap", "dining@nordsy.spb.ru", "AA77rr11!!"
            );

            if (inputStream == null) {
                return "Файл Excel не найден в почте";
            }

            // Чтение данных из Excel
            Map<String, List<String>> data = excelService.readExcel(inputStream);
            // Вывод данных
            data.forEach((key, value) -> {
                System.out.println(key + ": " + value);
            });
//            List<String[]> data = excelService.readExcel(inputStream);
//            for (String[] array : data) {
//                for (String value : array) {
//                    System.out.print(value + " "); // Вывод каждого элемента массива
//                }
//                System.out.println(); // Переход на новую строку после массива
//            }


             // Создание нового Excel
//            excelCreationService.createExcel(data, "processed_data.xlsx");

            return "Файл успешно обработан и сохранён как processed_data.xlsx";
        } catch (Exception e) {
            e.printStackTrace();
            return "Ошибка при обработке файла: " + e.getMessage();
        }
    }

}

