package utl;

import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.regex.Pattern;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;

/**
 * describe:
 *
 * @author phantom
 * @date 2019/06/24
 */
public class CommonMethods {

    /**
     * 判断是否为整数
     * @param str 字符串
     * @return boolean
     */
    public static boolean isInteger(String str) {
        Pattern pattern = Pattern.compile("^[-\\+]?[\\d]*$");
        return pattern.matcher(str).matches();
    }

    /**
     * 判断是否为浮点型
     * @param str 字符串
     * @return boolean
     */
    public static boolean isDouble(String str) {
        if (null == str || "".equals(str)) {
            return false;
        }
        Pattern pattern = Pattern.compile("^[-\\+]?[.\\d]*$");
        return pattern.matcher(str).matches();
    }

    public static String parseDate(Cell cell) {
        int format = cell.getCellStyle().getDataFormat();
        if (format == 14 || format == 31 || format == 57 || format == 58
                || (176<=format && format<=178) || (182<=format && format<=196)
                || (210<=format && format<=213) || (208==format ) ) {
            SimpleDateFormat simpleDateFormat = new SimpleDateFormat("yyyy-MM");
            double value = cell.getNumericCellValue();
            Date date = org.apache.poi.ss.usermodel.DateUtil.getJavaDate(value);
            return simpleDateFormat.format(date);
        }else {
            return cell.toString();
        }
    }
}
