package utl;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;

import static java.io.File.separator;

/**
 * describe:
 *
 * @author phantom
 * @date 2019/06/24
 */
public class GetWorkBook {

    public static Workbook getWorkBook(String filePath){
        Workbook wb = null;
        String suffix = filePath.substring(filePath.lastIndexOf("."));
        try {
            InputStream is = new FileInputStream(filePath);
            if (suffix.equals(".xls")){
                wb = new HSSFWorkbook(is);
            } else if (suffix.equals(".xlsx")){
                wb = new XSSFWorkbook(is);
            }
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
        return wb;
    }

    public static Workbook getWorkBook(String year, String organization, String fileName){
        //获取存在统战工作数据表的单位
        String parentDir = new GetYears().getPath();
        parentDir = parentDir + separator + year + separator + organization;
        String filePath = parentDir + separator + fileName;
        Workbook wb = null;
        String suffix = filePath.substring(filePath.lastIndexOf("."));
        try {
            InputStream is = new FileInputStream(filePath);
            if (suffix.equals(".xls")){
                wb = new HSSFWorkbook(is);
            } else if (suffix.equals(".xlsx")){
                wb = new XSSFWorkbook(is);
            }
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
        return wb;
    }
}
