package com.myapp.exelfromemail.service;

import jakarta.mail.*;
import jakarta.mail.search.FlagTerm;
import org.apache.poi.ss.usermodel.*;
import org.springframework.stereotype.Service;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Map;
import java.util.Properties;

@Service
public class ExcelReaderService {

    public File downloadExcel(String folderName, String subjectKeyword) throws Exception {
        Properties properties = new Properties();
        properties.put("mail.store.protocol", "imap");
        properties.put("mail.imap.host", "cgp.nordsy.spb.ru");
        properties.put("mail.imap.port", "143");
        properties.put("mail.imap.ssl.enable", "false");

        Store store = null;
        Folder folder = null;

        try {
            // Подключение к хранилищу
            Session session = Session.getDefaultInstance(properties, null);
            store = session.getStore("imap");
            store.connect("cgp.nordsy.spb.ru", "10390@nordsy.spb.ru", "rfnfvfhfy342A");

            // Открытие папки
            folder = store.getFolder(folderName);
            folder.open(Folder.READ_WRITE);

            // Поиск непрочитанных сообщений
            FlagTerm unreadFlag = new FlagTerm(new Flags(Flags.Flag.SEEN), false);
            Message[] messages = folder.search(unreadFlag);

            for (Message message : messages) {
                if (message.getSubject() != null && message.getSubject().contains(subjectKeyword)) {
                    Object content = message.getContent();

                    if (content instanceof Multipart multipart) {
                        for (int i = 0; i < multipart.getCount(); i++) {
                            BodyPart bodyPart = multipart.getBodyPart(i);

                            if (Part.ATTACHMENT.equalsIgnoreCase(bodyPart.getDisposition()) &&
                                    bodyPart.getFileName().endsWith(".xlsx")) {

                                // Сохранение вложения
                                File file = new File(System.getProperty("java.io.tmpdir"), bodyPart.getFileName());
//                                File file = new File("C:/MyFolder", bodyPart.getFileName());
//                                File file = new File("/home/user/MyFolder", bodyPart.getFileName());


                                try (InputStream inputStream = bodyPart.getInputStream();
                                     FileOutputStream outputStream = new FileOutputStream(file)) {

                                    byte[] buffer = new byte[1024];
                                    int bytesRead;
                                    while ((bytesRead = inputStream.read(buffer)) != -1) {
                                        outputStream.write(buffer, 0, bytesRead);
                                    }
                                }

                                // Пометка сообщения как прочитанного
                                message.setFlag(Flags.Flag.SEEN, true);
                                return file;
                            }
                        }
                    }
                }
            }
            throw new Exception("Excel файл не найден");
        } finally {
            if (folder != null && folder.isOpen()) {
                folder.close(false);
            }
            if (store != null && store.isConnected()) {
                store.close();
            }
        }
    }
}

