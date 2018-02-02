package test;

import com.google.common.collect.ImmutableList;
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

    public static final List<String> headers = ImmutableList.of("Адрес электронной почты", "Фамилия", "Имя", "Отчество", "Дата рождения",
            "Номер телефона", "Регион проживания либо город", "Гражданство",
            "Год окончания вуза (последнего места обучения)", "Наименование вуза",
            "Образование степень", "Специальность или факультет");

    public static final String PLACE = "Проживает: ";
    public static final String CITIZENSHIP = "Гражданство: ";
    public static final String EDUCATION = "Образование";

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
                                }
                            }
                        }
                    }
                }
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

    public static void createExcel( List<List<String>> data, String outDir){
        final String FILE_NAME = outDir + "\\Summary table.xlsx";
        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet();

        int rowNum = 0;
        System.out.println("Creating excel");

        for (List<String> datatype : data) {
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

    private static List<List<String>> processDataForExcel(Map<Integer, List<String>> contentMap) {
        List<List<String>> dataForExcel = new ArrayList<List<String>>();
        dataForExcel.add(headers);
        for (int i = 0;contentMap.size()>i ; i++){
            List<String> content = contentMap.get(i);
            List<String> rowsAfterHeader = getRowsFromContent(content);
            dataForExcel.add(rowsAfterHeader);
        }
        return dataForExcel;
    }

    private static List<String> getRowsFromContent(List<String> content) {
        List<String> rows = new ArrayList<String>();
        String [] fio = content.get(0).split(" ");
        String family = fio[0];
        String name = fio[1];
        String surname = fio[2];
        String dataR = content.get(2);
        String email = getEmail(content);
        String telephone = getPhone(content);
        String place = getPlace(content);
        String citizenShip = getCitizenShip(content);
        String[] education = getEducation(content);
        String eduType = education[0];
        String eduYear = education[1];
        String eduName = education[2];
        String eduSpec = education[3];

        rows.add(email);
        rows.add(family);
        rows.add(name);
        rows.add(surname);
        rows.add(dataR);
        rows.add(telephone);
        rows.add(place);
        rows.add(citizenShip);
        rows.add(eduYear);
        rows.add(eduName);
        rows.add(eduType);
        rows.add(eduSpec);

        return rows;
    }

    private static String[] getEducation(List<String> content) {
        String [] str = new String[]{};
        for (int i = 0; content.size() > i ; i++)
        {
            if (content.get(i) != null && content.get(i).equals(EDUCATION))
            {
                return new String[]{content.get(i+1), content.get(i+2), content.get(i+3), content.get(i+4)};
            }
        }
        return str;
    }

    private static String getCitizenShip(List<String> content) {
        String str = "";
        for (int i = 0; content.size() > i ; i++)
        {
            if (content.get(i) != null && content.get(i).equals(CITIZENSHIP))
            {
                return content.get(i+1);
            }
        }
        return str;
    }

    private static String getPlace(List<String> content) {
        String place = "";
        for (int i = 0; content.size() > i ; i++)
        {
            if (content.get(i) != null && content.get(i).equals(PLACE))
            {
                return content.get(i+1);
            }
        }
        return place;
    }

    private static String getPhone(List<String> content) {
        String phone = "";
        Pattern p = Pattern.compile("^\\+\\d\\S\\(\\d\\d\\d\\)\\S\\d\\d\\d\\d\\d\\d\\d$", Pattern.CASE_INSENSITIVE);
        for (String str : content)
        {
            String result = getStrFromPattern(p, str);
            if (result != null) {
                return result;
            }
        }
        return phone;
    }

    private static String getEmail(List<String> content) {
        String email = "";
        Pattern p = Pattern.compile("\\b[A-Z0-9._%+-]+@[A-Z0-9.-]+\\.[A-Z]{2,4}\\b", Pattern.CASE_INSENSITIVE);
        for (String str : content)
        {
            String result = getStrFromPattern(p, str);
            if (result != null) {
                return result;
            }
        }
        return email;
    }

    private static String getStrFromPattern(Pattern p, String str) {
        if ( str != null && str != " ") {
            Matcher matcher = p.matcher(str);
            if (matcher.matches()) {
                return matcher.group();
            }
        }
        return null;
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
