package test;

import javax.swing.*;
import java.awt.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;

/**
 * @author Andrey Smirnov
 * @date 01.02.2018
 */
public class GuiCreator {
    public static void guiApp() {

        //Main frame
        final JFrame guiFrame = new JFrame();
        guiFrame.setDefaultCloseOperation(WindowConstants.EXIT_ON_CLOSE);
        guiFrame.setTitle("Парсер резюме");
        guiFrame.setSize(320, 250);
        guiFrame.setLocationRelativeTo(null);

        //Labels
        JLabel inDirLbl = new JLabel("Исходная директория :");
        final JLabel inDirectory = new JLabel("");
        JLabel outDirLbl = new JLabel("Куда сохраняем :");
        final JLabel outDirectory = new JLabel("");
        JPanel panel = new JPanel();
        panel.setLayout(new BoxLayout(panel, BoxLayout.Y_AXIS));
        panel.add(inDirLbl);
        panel.add(inDirectory);
        panel.add(outDirLbl);
        panel.add(outDirectory);

        //Buttons
        JButton choseBtn = new JButton("1.Исходная директория");
        JButton whereBtn = new JButton("2.Куда сохраняем");
        JButton parseBtn = new JButton("3.Начать парсинг");
        JPanel btnPanel = new JPanel();
        btnPanel.setLayout(new GridLayout(3, 1));
        btnPanel.add(choseBtn);
        btnPanel.add(whereBtn);
        btnPanel.add(parseBtn);
        JPanel east = new JPanel(new GridBagLayout());
        GridBagConstraints gbc = new GridBagConstraints();
        gbc.anchor = GridBagConstraints.CENTER;
        gbc.weighty = 1;
        east.add(btnPanel, gbc);

        //Button listeners
        choseBtn.addActionListener(new ActionListener() {
            JFileChooser chooser;
            public void actionPerformed(ActionEvent event) {
                chooser = new JFileChooser();
                chooser.setCurrentDirectory(new java.io.File("."));
                chooser.setDialogTitle("Выберите директорию с резюме");
                chooser.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
                chooser.setAcceptAllFileFilterUsed(false);
                if (chooser.showOpenDialog(chooser) == JFileChooser.APPROVE_OPTION) {
                    System.out.println("getCurrentDirectory(): "
                            + chooser.getCurrentDirectory());
                    inDirectory.setText(chooser.getSelectedFile().getPath());
                    System.out.println("getSelectedFile() : "
                            + chooser.getSelectedFile());
                } else {
                    System.out.println("No Selection ");
                }
            }
        });
        whereBtn.addActionListener(new ActionListener() {
            JFileChooser chooser;
            public void actionPerformed(ActionEvent event) {
                chooser = new JFileChooser();
                chooser.setCurrentDirectory(new java.io.File("."));
                chooser.setDialogTitle("Выберите директорию с резюме");
                chooser.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
                chooser.setAcceptAllFileFilterUsed(false);
                if (chooser.showOpenDialog(chooser) == JFileChooser.APPROVE_OPTION) {
                    System.out.println("getCurrentDirectory(): "
                            + chooser.getCurrentDirectory());
                    outDirectory.setText(chooser.getSelectedFile().getPath());
                    System.out.println("getSelectedFile() : "
                            + chooser.getSelectedFile());
                } else {
                    System.out.println("No Selection ");
                }
            }
        });
        parseBtn.addActionListener(new ActionListener() {
            public void actionPerformed(ActionEvent event) {
                if (inDirectory.getText().length() > 1 && outDirectory.getText().length() > 1) {
                    Parser.startParsing(inDirectory.getText(), outDirectory.getText(), guiFrame);
                } else JOptionPane.showMessageDialog(guiFrame, "Директория не выбрана. Убедитесь что выбрали обе директории.",
                        "Предупреждение", JOptionPane.WARNING_MESSAGE);
            }
        });

        //The JFrame uses the BorderLayout layout manager.
        guiFrame.add(panel, BorderLayout.NORTH);
        guiFrame.add(east, BorderLayout.CENTER);

        //make sure the JFrame is visible
        guiFrame.setVisible(true);
    }
}
