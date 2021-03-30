import org.jopendocument.dom.OOUtils;
import org.jopendocument.dom.spreadsheet.Sheet;
import org.jopendocument.dom.spreadsheet.SpreadSheet;

import javax.swing.table.DefaultTableModel;
import javax.swing.table.TableModel;
import java.io.File;
import java.io.IOException;

public class Main {
    private final static String inputDir = "/home/jhaugh/teaching/cs351/spring2021/assignments/project2/documents/grades/";
    private final static String outputDir = "/home/jhaugh/teaching/cs351/spring2021/assignments/project2/documents/grades/students/";
    public static void main(String[] args) throws IOException {
        // Load the file.
        File file = new File(inputDir + "CS351_Project2.ods");
        final SpreadSheet sp = SpreadSheet.createFromFile(file);

        for (int i = 2; i < sp.getSheetCount(); i++) {
            final SpreadSheet sp_copy = SpreadSheet.createFromFile(file);
            removeOtherTabs(sp_copy, i);
            String outputFileName = outputDir + sp.getSheet(i).getName() + ".ods";
            System.out.println("Processing " + outputFileName);
            File outputFile = new File(outputFileName);
            OOUtils.open(sp_copy.saveAs(outputFile));

            final String cmd = "libreoffice --headless --convert-to pdf --outdir " + outputDir + " " + outputFileName;

            try {
                System.out.println("Creating PDF for " + outputFileName);
                Process process = Runtime.getRuntime().exec(cmd);
                process.waitFor();
                System.out.println("PDF ready for " + outputFileName);
            }
            catch (InterruptedException exc) {
                System.out.println("Process was interrupted!");
            }
        }
    }

    private static void removeOtherTabs(SpreadSheet sp, int currentTab) {
        int sheet_count = sp.getSheetCount();
        for (int i = sheet_count - 1; i >= 0; i--) {
            if (i != currentTab) {
                sp.getSheet(i).detach();
            }
        }
    }
}
