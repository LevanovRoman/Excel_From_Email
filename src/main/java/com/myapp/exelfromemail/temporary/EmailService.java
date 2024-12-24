package com.myapp.exelfromemail.temporary;

import jakarta.mail.*;
import jakarta.mail.internet.MimeBodyPart;
import jakarta.mail.search.FlagTerm;
import org.springframework.stereotype.Service;
import jakarta.mail.Flags;


import java.io.File;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.util.Properties;

@Service
public class EmailService {

    public InputStream getExcelAttachment(String host, String storeType, String user, String password) throws Exception {
        // Настройка почтового соединения
        Properties properties = new Properties();
        properties.put("mail.store.protocol", "imap");
        properties.put("mail.imap.host", host);
        properties.put("mail.imap.port", "143");
        properties.put("mail.imap.ssl.enable", "false");

        Session emailSession = Session.getDefaultInstance(properties);
        Store store = emailSession.getStore("imap");
        store.connect(host, user, password);

        // Открываем папку "Входящие"
        Folder folder = store.getFolder("INBOX");
        folder.open(Folder.READ_ONLY);

        // Проверяем письма
        for (Message message : folder.getMessages()) {
            if (message.getContentType().contains("multipart")) {
                Multipart multipart = (Multipart) message.getContent();
                for (int i = 0; i < multipart.getCount(); i++) {
                    MimeBodyPart bodyPart = (MimeBodyPart) multipart.getBodyPart(i);
                    if (Part.ATTACHMENT.equalsIgnoreCase(bodyPart.getDisposition())) {
                        if (bodyPart.getFileName().endsWith(".xlsx")) {
                            return bodyPart.getInputStream();
                        }
                    }
                }
            }
        }
        return null;
    }

    public File downloadExcelAttachment(String folderName, String subjectKeyword) throws Exception {
        // Настройка IMAP-сессии
        Properties properties = new Properties();
        properties.put("mail.store.protocol", "imap");
        properties.put("mail.imap.host", "cgp.nordsy.spb.ru");
        properties.put("mail.imap.port", "143");
        properties.put("mail.imap.ssl.enable", "false");

        Session session = Session.getDefaultInstance(properties, null);
        Store store = session.getStore("imap");
        store.connect("cgp.nordsy.spb.ru", "10390@nordsy.spb.ru", "rfnfvfhfy342A");

        Folder folder = store.getFolder(folderName);
        folder.open(Folder.READ_WRITE);

        // Поиск непрочитанных писем с указанным ключевым словом в теме
        Message[] messages = folder.search(new FlagTerm(new Flags(Flags.Flag.SEEN), false));

        for (Message message : messages) {
            if (message.getSubject().contains(subjectKeyword)) {
                Multipart multipart = (Multipart) message.getContent();
                for (int i = 0; i < multipart.getCount(); i++) {
                    BodyPart bodyPart = multipart.getBodyPart(i);

                    if (Part.ATTACHMENT.equalsIgnoreCase(bodyPart.getDisposition())) {
                        String fileName = bodyPart.getFileName();
                        if (fileName.endsWith(".xlsx")) {
                            File file = new File(System.getProperty("java.io.tmpdir") + "/" + fileName);
                            try (FileOutputStream outputStream = new FileOutputStream(file)) {
                                ((MimeBodyPart) bodyPart).saveFile(file);
                            }
                            return file; // Возвращаем найденный файл
                        }
                    }
                }
            }
        }

        folder.close(false);
        store.close();

        throw new Exception("Excel файл не найден");
    }

    public File downloadExcelAttachment2(String folderName, String subjectKeyword) throws Exception {
        // Настройка IMAP-сессии
        Properties properties = new Properties();
        properties.put("mail.store.protocol", "imap");
        properties.put("mail.imap.host", "cgp.nordsy.spb.ru");
        properties.put("mail.imap.port", "143");
        properties.put("mail.imap.ssl.enable", "false");

        Session session = Session.getDefaultInstance(properties, null);
        Store store = session.getStore("imap");
        store.connect("cgp.nordsy.spb.ru", "10390@nordsy.spb.ru", "rfnfvfhfy342A");

        Folder folder = store.getFolder(folderName);
        folder.open(Folder.READ_WRITE);

        // Поиск непрочитанных писем
        Message[] messages = folder.getMessages();
        for (Message message : messages) {
            if (message.getSubject() != null && message.getSubject().contains(subjectKeyword)) {
                Object content = message.getContent();

                // Проверяем, является ли содержимое Multipart
                if (content instanceof Multipart) {
                    Multipart multipart = (Multipart) content;

                    for (int i = 0; i < multipart.getCount(); i++) {
                        BodyPart bodyPart = multipart.getBodyPart(i);

                        if (Part.ATTACHMENT.equalsIgnoreCase(bodyPart.getDisposition())) {
                            String fileName = bodyPart.getFileName();
                            if (fileName.endsWith(".xlsx")) {
                                File file = new File(System.getProperty("java.io.tmpdir") + "/" + fileName);
                                try (FileOutputStream outputStream = new FileOutputStream(file)) {
                                    ((MimeBodyPart) bodyPart).saveFile(file);
                                }
                                return file; // Возвращаем файл Excel
                            }
                        }
                    }
                } else if (content instanceof String) {
                    // Если письмо простое текстовое (для отладки)
                    System.out.println("Письмо содержит текст: " + content);
                }
            }
        }

        folder.close(false);
        store.close();

        throw new Exception("Excel файл не найден");
    }


//    public File downloadExcelAttachment3(String folderName, String subjectKeyword) throws Exception {
//        // Настройка IMAP-сессии
//        Properties properties = new Properties();
//        properties.put("mail.store.protocol", "imap");
//        properties.put("mail.imap.host", "cgp.nordsy.spb.ru");
//        properties.put("mail.imap.port", "143");
//        properties.put("mail.imap.ssl.enable", "false");
//
//        // Подключение к почтовому ящику
//        try (Store store = Session.getDefaultInstance(properties, null).getStore("imap")) {
//            store.connect("cgp.nordsy.spb.ru", "10390@nordsy.spb.ru", "rfnfvfhfy342A");
//            try (Folder folder = store.getFolder(folderName)) {
//                folder.open(Folder.READ_WRITE);
//
//                // Поиск писем
//                Message[] messages = folder.getMessages();
//                for (Message message : messages) {
//                    if (message.getSubject() != null && message.getSubject().contains(subjectKeyword)) {
//                        Object content = message.getContent();
//
//                        // Проверка на Multipart
//                        if (content instanceof Multipart multipart) {
//                            for (int i = 0; i < multipart.getCount(); i++) {
//                                BodyPart bodyPart = multipart.getBodyPart(i);
//
//                                if (Part.ATTACHMENT.equalsIgnoreCase(bodyPart.getDisposition()) &&
//                                        bodyPart.getFileName().endsWith(".xlsx")) {
//
//                                    // Создаем временный файл
//                                    File file = new File(System.getProperty("java.io.tmpdir"), bodyPart.getFileName());
//                                    try (InputStream inputStream = bodyPart.getInputStream();
//                                         FileOutputStream outputStream = new FileOutputStream(file)) {
//
//                                        // Копируем содержимое вложения в файл
//                                        byte[] buffer = new byte[1024];
//                                        int bytesRead;
//                                        while ((bytesRead = inputStream.read(buffer)) != -1) {
//                                            outputStream.write(buffer, 0, bytesRead);
//                                        }
//                                    }
//
//                                    return file; // Возвращаем файл
//                                }
//                            }
//                        } else if (content instanceof String) {
//                            // Для текстового письма
//                            System.out.println("Письмо содержит текст: " + content);
//                        }
//                    }
//                }
//
//                // Закрываем папку
//                folder.close(false);
//            }
//        }
//
//        throw new Exception("Excel файл не найден");
//    }

}

