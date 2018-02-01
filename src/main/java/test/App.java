package test;

/**
 * Hello world!
 */
public class App {
    public static void main(String[] args) {
        // запустить приложение
        javax.swing.SwingUtilities.invokeLater(new Runnable() {
            public void run() {
                GuiCreator.guiApp();
            }
        });
    }
}