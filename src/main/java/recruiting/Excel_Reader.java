package recruiting;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;

import java.io.FileInputStream;
import java.util.Iterator;
import java.util.Vector;

public class Excel_Reader {
    public static void main(String[] args) {
        String nameOfFile = "src\\main\\resources\\Candidates.xlsx";
        Vector dataSet = readExcelFile(nameOfFile);
        printCellDataToTheConsole(dataSet);
    }

    public static Vector readExcelFile(String nameOfFile) {
        Vector cellVectorHolder = new Vector();

        try {
            FileInputStream input = new FileInputStream(nameOfFile);
            POIFSFileSystem mySystem = new POIFSFileSystem(input);
            HSSFWorkbook workbook = new HSSFWorkbook(mySystem);
            HSSFSheet sheet = workbook.getSheetAt(0);
            Iterator iterRows = sheet.rowIterator();

            while (iterRows.hasNext()) {
                HSSFRow row = (HSSFRow) iterRows.next();
                Iterator iterCell = row.cellIterator();
                Vector storeVectorInCell = new Vector();
                while (iterCell.hasNext()) {
                    HSSFCell cell = (HSSFCell) iterCell.next();
                    storeVectorInCell.addElement(cell);
                }
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
        return cellVectorHolder;
    }

    private static void printCellDataToTheConsole(Vector dataSet) {
        try {
            for (int i = 0; i < 6000; i++) {

                Vector cellStoreVector = (Vector) dataSet.elementAt(i);
                for (int j = 0; j < cellStoreVector.size(); j++) {
                    HSSFCell myCell = (HSSFCell) cellStoreVector.elementAt(j);
                    String stringCellValue = myCell.toString();
                    System.out.print(j + " " + stringCellValue + "\t");
                }
                System.out.println();
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
