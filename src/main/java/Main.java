import org.apache.poi.xwpf.model.XWPFHeaderFooterPolicy;
import org.apache.poi.xwpf.usermodel.*;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

public class Main {
    public static void main(String args[]) throws IOException {

        XWPFDocument document = new XWPFDocument(); // creating a doc instants
        FileOutputStream out = new FileOutputStream(new File("C:\\Users\\RamzanLafir\\Desktop\\JAVA Plugin\\generated.docx")); // file location

        // 1st Paragraph
        XWPFParagraph paragraph = document.createParagraph();
        XWPFRun run = paragraph.createRun();
        run.setText("Generate document for word encode, created word document. Ramzan");

        //create table
        XWPFTable table = document.createTable();
        //create first row
        XWPFTableRow tableRowOne = table.getRow(0);
        tableRowOne.getCell(0).setText(" ID ");
        tableRowOne.addNewTableCell().setText(" Name ");
        tableRowOne.addNewTableCell().setText(" Location ");
        //create second row
        XWPFTableRow tableRowTwo = table.createRow();
        tableRowTwo.getCell(0).setText(" 01 ");
        tableRowTwo.getCell(1).setText(" Ramzan ");
        tableRowTwo.getCell(2).setText(" Negombo ");
        //create third row
        XWPFTableRow tableRowThree = table.createRow();
        tableRowThree.getCell(0).setText(" 02 ");
        tableRowThree.getCell(1).setText(" Tom ");
        tableRowThree.getCell(2).setText(" Colombo ");

        // 2nd Paragraph
        XWPFParagraph paragraphOneRunThree = document.createParagraph();
        XWPFRun run2 = paragraphOneRunThree.createRun();
        // paragraphOneRunThree.setStrike(true);
        run2.setFontSize(40);
        run2.setSubscript(VerticalAlign.SUBSCRIPT);
        run2.setText("Heading Title");
        document.write(out);
        out.close();

        System.out.println("Complete");

    }
}
