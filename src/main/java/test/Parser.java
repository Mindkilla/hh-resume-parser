package test;

import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.usermodel.*;

import javax.swing.*;
import java.io.*;
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * @author Andrey Smirnov
 * @date 01.02.2018
 */
class Parser {

    static void startParsing(String inDir, String outDir, JFrame guiFrame) {
        int counter = 0;
        Map<Integer, List<String>> contentMap = new HashMap<Integer, List<String>>();
        try {
            List<File> files = getFiles(new File(inDir));
            if (files.isEmpty()) {
                JOptionPane.showMessageDialog(guiFrame, "В директории нет файлов для обработки.", "Warning", JOptionPane
                        .WARNING_MESSAGE);
                return;
            }
            for (final File file : files) {
                FileInputStream fis = new FileInputStream(file.getAbsolutePath());
                XWPFDocument xdoc = new XWPFDocument(OPCPackage.open(fis));
                List<String> content = new LinkedList<String>();
                for (XWPFTable tbl : xdoc.getTables()) {
                    for (XWPFTableRow row : tbl.getRows()) {
                        for (XWPFTableCell cell : row.getTableCells()) {
                            for (XWPFParagraph p : cell.getParagraphs()) {
                                for (XWPFRun r : p.getRuns()) {
                                    System.out.println(r.getText(0));
                                    content.add(r.getText(0));
                                    /*String text = r.getText(0);
                                    if (text != null && text.contains("needle")) {
                                        text = text.replace("needle", "haystack");
                                        r.setText(text,0);
                                    }*/
                                }
                            }
                        }
                    }
                }
                /*Iterator<IBodyElement> bodyElementIterator = xdoc.getBodyElementsIterator();
                while (bodyElementIterator.hasNext()) {
                    IBodyElement element = bodyElementIterator.next();
                    if ("TABLE".equalsIgnoreCase(element.getElementType().name())) {
                        List<XWPFTable> tableList = element.getBody().getTables();
                        for (XWPFTable table : tableList) {
                            System.out.println("Total Number of Rows of Table:" + table.getNumberOfRows());
                            System.out.println(table.getText());
                        }
                        counter++;
                    }
                }*/
                contentMap.put(counter, content);
                counter++;
            }
        } catch (Exception ex) {
            JOptionPane.showMessageDialog(guiFrame, ex.getMessage(), "Ошибка", JOptionPane.ERROR_MESSAGE);
            System.out.println(ex.getMessage());
        }
        if (counter > 0) {
            createExcel(processDataForExcel(contentMap), outDir);
            JOptionPane.showMessageDialog(guiFrame, "Итого обработано файлов :" + counter, "INFO", JOptionPane.INFORMATION_MESSAGE);
        }
    }

    private static void createExcel(Map<Integer, List<String>> content, String outDir) {
        //template
    }

    public static void createExcel(Object[][] data, String outDir){
        final String FILE_NAME = outDir + "\\Summary table.xlsx";
        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet();

        int rowNum = 0;
        System.out.println("Creating excel");

        for (Object[] datatype : data) {
            Row row = sheet.createRow(rowNum++);
            int colNum = 0;
            for (Object field : datatype) {
                Cell cell = row.createCell(colNum++);
                CellStyle style = workbook.createCellStyle();
                style.setWrapText(true);
                style.setBorderBottom(BorderStyle.THIN);
                style.setBorderRight(BorderStyle.THIN);
                style.setBorderLeft(BorderStyle.THIN);
                style.setBorderTop(BorderStyle.THIN);
                cell.setCellStyle(style);
                if (field instanceof String) {
                    cell.setCellValue((String) field);
                } else if (field instanceof Integer) {
                    cell.setCellValue((Integer) field);
                }
            }
        }
        try {
            FileOutputStream outputStream = new FileOutputStream(FILE_NAME);
            workbook.write(outputStream);
            workbook.close();
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }

        System.out.println("Done");
    }

    private static Object [][] processDataForExcel(Map<Integer, List<String>> contentMap) {
        //template
        /*[Морозов Павел Андреевич, Мужчина, 21 год, родился , 23 апреля 1996, +7 (985) 9855831,
        , — предпочитаемый способ связи, mr.morozov96@gmail.com, Проживает: , Москва, , м. Войковская
        , Гражданство: , Россия, , есть разрешение на работу: Россия, Не готов к переезду
        , готов к редким командировкам, Желаемая должность и зарплата, Менеджер по подбору персонала
        , Начало карьеры, студенты, • Управление персоналом, • Административный персонал, • Продажи
        , Занятость: полная занятость, График работы: полный день, Желательное время в пути до работы: не имеет значения
        , 40 000, null, руб., Образование, Высшее, 2017, Московский психолого-социальный университет, Москва, юридический
        , Управление персоналом, Ключевые навыки, Знание языков, Русский , — родной, Английский , — базовые знания
        , Навыки, Управление командой,   , Подбор персонала,   , Оценка кандидатов,   , Адаптация персонала
        ,   , Аутплейсмент, Дополнительная информация, Обо мне, В последние годы проходил обучение без возможности работать.]
        */
        List<List<String>> data2 = new ArrayList<List<String>>();
        List<String> header = new ArrayList<String>();
        header.addAll(new String[]{"Адрес электронной почты", "Фамилия", "Имя", "Отчество", "Дата рождения",
                "Номер телефона", "Регион проживания либо город", "Гражданство",
                "Год окончания вуза (последнего места обучения)", "Наименование вуза",
                "Образование степень", "Специальность или факультет"})
        data2.add(header); = new String[]{"Адрес электронной почты", "Фамилия", "Имя", "Отчество", "Дата рождения",
                "Номер телефона", "Регион проживания либо город", "Гражданство",
                "Год окончания вуза (последнего места обучения)", "Наименование вуза",
                "Образование степень", "Специальность или факультет"};


        for (int i = 0;contentMap.size()>i ; i++){
            List<String> content = contentMap.get(i);
            data2[i+1] =  new String[]{content.get(0)};

        }
        /*String fio = contentMap.get(0).get(0);
        String dataR = contentMap.get(0).get(2);
        String email = contentMap.get(0).get(5);
        Pattern p = Pattern.compile("\\b[A-Z0-9._%+-]+@[A-Z0-9.-]+\\.[A-Z]{2,4}\\b", Pattern.CASE_INSENSITIVE);
        for (String str : contentMap.get(0))
        {
            Matcher matcher = p.matcher(str);
            if (matcher.matches()) {
                String email2 = matcher.group();
            }
        }*/
        return data2;
    }

    private static List<File> getFiles(File inFolder) {
        List<File> fileList = new ArrayList<File>();
        for (final File file : inFolder.listFiles()) {
            if (file.isFile()) {
                fileList.add(file);
            }
        }
        return fileList;
    }
}
