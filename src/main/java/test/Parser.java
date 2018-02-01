package test;

import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xwpf.usermodel.*;

import javax.swing.*;
import java.io.*;
import java.util.*;

/**
 * @author Andrey Smirnov
 * @date 01.02.2018
 */
public class Parser {

    public static void startParsing(String inDir, String outDir, JFrame guiFrame) {
        int counter = 0;
        Map<Integer, List<String>> contentMapa = new HashMap<Integer, List<String>>();
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
                contentMapa.put(counter, content);
                counter++;
            }
        } catch (Exception ex) {
            JOptionPane.showMessageDialog(guiFrame, ex.getMessage(), "Ошибка", JOptionPane.ERROR_MESSAGE);
            System.out.println(ex.getMessage());
        }
        createExcel(processDataForExcel(contentMapa));
        if (counter > 0) {
            JOptionPane.showMessageDialog(guiFrame, "Итого обработано файлов :" + counter, "INFO", JOptionPane.INFORMATION_MESSAGE);
        }
    }

    private static void createExcel(Map<Integer, List<String>> content) {
        //template
    }

    private static Map processDataForExcel(Map<Integer, List<String>> contentMapa) {
        //template
        String fio = contentMapa.get(0).get(0);
        return new HashMap();
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
