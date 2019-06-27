package utl;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import static java.io.File.separator;

/**
 * describe:
 *
 * @author phantom
 * @date 2019/06/26
 */
public class WriteTableToExcel {
    private String path;

    public WriteTableToExcel() {
        this.path = System.getProperty("user.dir");
    }

    public void write(String name, String[][] data){
        String filePath = path + separator + name + ".xlsx";
        File file = new File(filePath);
        if (file.exists()){
            file.delete();
        }
        Workbook workbook = new SXSSFWorkbook();
        Sheet sheet = workbook.createSheet();
        for (int i = 0; i < data.length; i++) {
            Row row = sheet.createRow(i);
            row.setHeightInPoints(30);
            for (int j = 0; j < data[i].length; j++) {
                Cell cell = row.createCell(j);
                cell.setCellValue(data[i][j]);
            }
        }
        FileOutputStream fos = null;
        try {
            fos  = new FileOutputStream(new File(filePath));
            workbook.write(fos);
            fos.close();
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }



}
