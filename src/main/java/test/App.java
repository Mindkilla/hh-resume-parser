package test;

public class App {
    public static void main(String[] args) {
        // запустить приложение
        javax.swing.SwingUtilities.invokeLater(new Runnable() {
            public void run() {
                AppGui.createWindow();
            }
        });
    }
}