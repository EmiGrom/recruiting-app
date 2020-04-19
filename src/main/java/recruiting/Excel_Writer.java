package recruiting;

import java.io.File;

import jxl.Workbook;
import jxl.write.*;
import jxl.write.Label;

public class Excel_Writer {
    public static void main(String[] args) {
        WritableWorkbook workbook;
        try {
            workbook = Workbook.createWorkbook(new File("src\\main\\resources\\Candidates.xlsx"));

            WritableSheet writableSheet = workbook.createSheet("Sheet no 1", 0);
            Label label = new Label(0, 2, "First candidate");
            Label label2 = new Label(0, 3, "Second candidate");
            Label label3 = new Label(0, 4, "Third candidate");
            Label label4 = new Label(0, 5, "Fourth candidate");

            Label label5 = new Label(1, 2, "Java");
            Label label6 = new Label(1, 3, "C#");
            Label label7 = new Label(1, 4, "PHP");
            Label label8 = new Label(1, 5, "Python");

            Label label9 = new Label(2, 2, "5");
            Label label10 = new Label(2, 3, "7");
            Label label11 = new Label(2, 4, "3");
            Label label12 = new Label(2, 5, "8");

            Label label13 = new Label(3, 2, "multicultural environment");
            Label label14 = new Label(3, 3, "corporation");
            Label label15 = new Label(3, 4, "benefits");
            Label label16 = new Label(3, 5, "start-up");
            writableSheet.addCell(label);
            writableSheet.addCell(label2);
            writableSheet.addCell(label3);
            writableSheet.addCell(label4);
            writableSheet.addCell(label5);
            writableSheet.addCell(label6);
            writableSheet.addCell(label7);
            writableSheet.addCell(label8);
            writableSheet.addCell(label9);
            writableSheet.addCell(label10);
            writableSheet.addCell(label11);
            writableSheet.addCell(label12);
            writableSheet.addCell(label13);
            writableSheet.addCell(label14);
            writableSheet.addCell(label15);
            writableSheet.addCell(label16);
            int i = 0;
            int j = 1;

            label = new Label(i++, j, "Candidates");
            writableSheet.addCell(label);
            label = new Label(i++, j, "Technologies");
            writableSheet.addCell(label);
            label = new Label(i++, j, "Years of experience");
            writableSheet.addCell(label);
            label = new Label(i++, j, "Expectations");
            writableSheet.addCell(label);
            j++;

            label = new Label(i++, j, "Reserve list of candidates");
            writableSheet.addCell(label);
            label = new Label(i++, j, "Technologies");
            writableSheet.addCell(label);
            label = new Label(i++, j, "Years of experience");
            writableSheet.addCell(label);
            label = new Label(i++, j, "Expectations");
            writableSheet.addCell(label);

            workbook.write();
            workbook.close();
            System.out.println("Creating excel");
            System.out.println("Finish");

        } catch (Exception e) {
            System.out.println(e);
        }
    }

}
