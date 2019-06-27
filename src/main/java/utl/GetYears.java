package utl;


import java.io.File;
import java.util.Arrays;

import static java.io.File.separator;

/**
 * describe: 主要通过解析“北京市高校党员人员信息”文件获得可以统一的年份信息
 * @author phantom
 * @date 2019/06/24
 */
public class GetYears {
    private String path;
    private String[] years;

    public GetYears() {
        this.path = System.getProperty("user.dir") + separator + "北京市高校党外人员信息";
        parseDir();
    }

    private void parseDir(){
        File dir = new File(path);
        String[] files = dir.list();
        this.years = new String[files.length];
        System.arraycopy(files, 0, years, 0, files.length);
    }

    public String[] getYears() {
        return years;
    }

    public String getPath() {
        return path;
    }

    public static void main(String[] args) {
        GetYears getYears = new GetYears();
        System.out.println(Arrays.toString(getYears.years));
    }
}
