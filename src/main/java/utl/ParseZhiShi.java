package utl;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.File;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

import static java.io.File.separator;

/**
 * describe:
 *
 * @author phantom
 * @date 2019/06/25
 */
public class ParseZhiShi {
    private String year;
    private String organization;
    private String[][] data;

    public ParseZhiShi(String year, String organization) {
        this.year = year;
        this.organization = organization;
        if (organization.equals("所有学校")){
            parseAllExcel();
        }else {
            parseOneExcel();
        }
    }

    private void parseOneExcel() {
        //创建存放数据的列表
        List<Bin> dataList = new ArrayList<Bin>();
        //获取存在统战工作数据表的单位
        String parentDir = new GetYears().getPath() + separator +
                this.year + separator + organization;
        //获得操作的对象
        Workbook wb = GetWorkBook.getWorkBook(this.year, this.organization, getTargetFileName(parentDir));
        //以上代码是为了获取每一个目标文件的操作对象，接下来读取每一个文件中的内容，存入文件中
        Sheet tempSheet = wb.getSheetAt(0);
        //获取统战工作表中需要统计的信息所在的行
        for (int i = 5; i < 11; i++) {
            Row temp = tempSheet.getRow(i);
            Bin bin = new Bin();
            for (int j = 0; j < Constant.ZHISHI; j++) {
                if (j >= 1){
                    if (temp.getCell(j).equals(null) || temp.getCell(j).toString().equals("")){
                        bin.add(String.valueOf(0));
                        continue;
                    }else {
                        if (temp.getCell(j).getCellType() == Cell.CELL_TYPE_NUMERIC){
                            bin.add(String.valueOf((int) temp.getCell(j).getNumericCellValue()));
                        }else {
                            bin.add(temp.getCell(j).toString());
                        }
                    }
                }else {
                    if (temp.getCell(j).equals(null) || temp.getCell(j).toString().equals("")){
                        bin.add("");
                        continue;
                    }else {
                    bin.add(temp.getCell(j).toString());
                    }
                }
            }
            dataList.add(bin);
        }
        //将datalist中的数据存入data中
        updateData(dataList);
    }

    private void parseAllExcel() {
        //获取存在统战工作数据表的单位
        List<String> organizations = new ArrayList<>();
        String parentDir = new GetYears().getPath();
        parentDir = parentDir + separator + this.year;
        File organizationsDir = new File(parentDir);
        String[] organizationNames = organizationsDir.list();
        for (String item: organizationNames) {
            String tempDir = parentDir + separator + item;
            File tempOrganizationDir = new File(tempDir);
            String[] fileNames = tempOrganizationDir.list();
            for (String excelName: fileNames) {
                if (excelName.contains("知识分子")){
                    organizations.add(item);
                    break;
                }
            }
        }
        List<Bin> dataList = new ArrayList<Bin>();
        //初始化6个Bin使得每个Bin的List中存放的是各个表格的累积的和
        for (int i = 0; i < 6; i++) {
            Bin bin = new Bin();
            dataList.add(bin);
        }
        
        int count = 0; //计数器
        //逐个读取包含“统战工作”的单位的表格信息
        for (String organization: organizations) {
            count++;
            String tempOrganizationDir = parentDir + separator + organization;

            String fileName = getTargetFileName(tempOrganizationDir);

            String filePath = tempOrganizationDir + separator + fileName;

            Workbook wb = GetWorkBook.getWorkBook(filePath);

            //以上代码是为了获取每一个目标文件的操作对象，接下来读取每一个文件中的内容，存入文件中
            Sheet tempSheet = wb.getSheetAt(0);
            //获取统战工作表中需要统计的信息所在的行
            for (int i = 5; i < 11; i++) {
                Row temp = tempSheet.getRow(i);
                if (count == 1){
                    for (int j = 0; j < Constant.ZHISHI; j++) {
                        if (j >= 1){
                            if (temp.getCell(j).equals(null) || temp.getCell(j).toString().equals("")){
                                dataList.get(i-5).add(String.valueOf(0));
                                continue;
                            }else {
                                if (temp.getCell(j).getCellType() == Cell.CELL_TYPE_NUMERIC){
                                    dataList.get(i-5).add(String.valueOf((int) temp.getCell(j).getNumericCellValue()));
                                }else {
                                    dataList.get(i-5).add(temp.getCell(j).toString());
                                }
                            }
                        }else {
                            if (temp.getCell(j).equals(null) || temp.getCell(j).toString().equals("")){
                                dataList.get(i-5).add("");
                                continue;
                            }else {
                                dataList.get(i-5).add(temp.getCell(j).toString());
                            }
                        }
                    }
                }else {
                    for (int j = 0; j < Constant.ZHISHI; j++) {
                        if (j == 0){
                            continue;
                        }else {
                            String content1 = temp.getCell(j).toString();
                            String content2 = dataList.get(i - 5).get(j);
                            double value1 = Double.valueOf(content1);
                            double value2 = Double.valueOf(content2);
                            int value = (int) value1 + (int) value2;
                            dataList.get(i - 5).changeValue(j, String.valueOf(value));
                        }
                    }
                }
            }
        }
        //将datalist中的数据存入data中
        updateData(dataList);

    }

    /**
     * 给出父目录获取指定的文件名字
     * @param parentDir 父目录
     * @return 指定的文件的名字
     */
    private String getTargetFileName(String parentDir) {
        String fileName = "";
        String[] tempNames = new File(parentDir).list();
        for (String name: tempNames) {
            if (name.contains("知识分子")){
                fileName = name;
                break;
            }
        }
        return fileName;
    }

    public String[][] getData() {
        return data;
    }

    /**
     * 将列表中的数据存入data中
     * @param dataList 列表中的数据
     */
    private void updateData(List<Bin> dataList) {
        data = new String[dataList.size()][Constant.ZHISHI];
        for (int i = 0; i < data.length; i++) {
            for (int j = 0; j < Constant.ZHISHI; j++) {
                data[i][j] = dataList.get(i).get(j);
            }
        }
    }


    public static void main(String[] args) {
        ParseZhiShi parseZhiShi = new ParseZhiShi("2019", "所有学校");
        String[][] data = parseZhiShi.getData();
        for (int i = 0; i < data.length; i++) {
            for (int j = 0; j < data[0].length; j++) {
                System.out.print(data[i][j] + ", ");
            }
            System.out.println();
        }
    }
}
