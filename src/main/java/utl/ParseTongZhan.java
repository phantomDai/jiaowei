package utl;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.*;
import java.util.ArrayList;
import java.util.List;

import static java.io.File.separator;
import static utl.CommonMethods.isDouble;
import static utl.CommonMethods.isInteger;

/**
 * describe:
 * 解析统战信息表格
 * @author phantom
 * @date 2019/06/24
 */
public class ParseTongZhan {
    private String year;
    private String organization;
    private String[][] data;

    public ParseTongZhan(String year, String organization) {
        this.year = year;
        this.organization = organization;
        if (organization.equals("所有学校")){
            parseAllExcel();
        }else {
            parseOneExcel();
        }

    }

    private void parseOneExcel(){
        //创建存放数据的列表
        List<Bin> dataList = new ArrayList<Bin>();
        //获取存在统战工作数据表的单位
        String parentDir = new GetYears().getPath();
        parentDir = parentDir + separator + this.year + separator + organization;
        String fileName = "";
        fileName = getTargetFileName(parentDir);

        String filePath = parentDir + separator + fileName;
        //获得操作的对象
        Workbook wb = GetWorkBook.getWorkBook(filePath);


        //以上代码是为了获取每一个目标文件的操作对象，接下来读取每一个文件中的内容，存入文件中
        Sheet tempSheet = wb.getSheetAt(0);
        //获取统战工作表中需要统计的信息所在的行
        int numberOfRows = tempSheet.getPhysicalNumberOfRows();
        int count = 0;

        for (int i = 0; i < numberOfRows; i++) {
            Row temp = tempSheet.getRow(i);
            String content = temp.getCell(0).toString();
            if (content.contains("注")){
                break;
            }
            if (!(isInteger(content) || isDouble(content))){
                continue;
            }else {
                String content1 = temp.getCell(1).toString();
                if ( content1.equals("")){
                    continue;
                }
                Bin bin = new Bin();
                count++;
                Row tempRow = tempSheet.getRow(i);
                for (int j = 0; j < Constant.TONGZHANCOL; j++) {
                    if (j == 0){
                        bin.add(String.valueOf(count));
                    }else if (j >= 3){
                        if (tempRow.getCell(j).equals(null)){
                            continue;
                        }
                        if (tempRow.getCell(j).getCellType() == Cell.CELL_TYPE_NUMERIC){
                            bin.add(String.valueOf((int) tempRow.getCell(j).getNumericCellValue()));
                        }else {
                            bin.add(tempRow.getCell(j).toString());
                        }
                    }else {
                        bin.add(tempRow.getCell(j).toString());
                    }
                }
                dataList.add(bin);
            }
        }
        //将datalist中的数据存入data中
        data = new String[dataList.size() + 1][Constant.TONGZHANCOL];
        for (int i = 1; i < data.length; i++) {
            for (int j = 0; j < dataList.get(i - 1).size(); j++) {
                data[i][j] = dataList.get(i - 1).get(j);
            }
        }
        String[] tempArray = calData(dataList.size());
        for (int i = 0; i < Constant.TONGZHANCOL; i++) {
            data[0][i] = tempArray[i];
        }
    }

    private void parseAllExcel(){
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
                if (excelName.contains("统战工作")){
                    organizations.add(item);
                    break;
                }
            }
        }
        List<Bin> dataList = new ArrayList<Bin>();
        int count = 0; //计数器
        //逐个读取包含“统战工作”的单位的表格信息
        for (String organization: organizations) {
            String tempOrganizationDir = parentDir + separator + organization;

            String fileName = getTargetFileName(tempOrganizationDir);

            String filePath = tempOrganizationDir + separator + fileName;

            Workbook wb = GetWorkBook.getWorkBook(filePath);

            //以上代码是为了获取每一个目标文件的操作对象，接下来读取每一个文件中的内容，存入文件中
            Sheet tempSheet = wb.getSheetAt(0);
            int numberOfRows = tempSheet.getPhysicalNumberOfRows();
            for (int i = 0; i < numberOfRows; i++) {
                Row temp = tempSheet.getRow(i);
                String content = temp.getCell(0).toString();
                if (content.contains("注")){
                    break;
                }
                if (!(isInteger(content) || isDouble(content))){
                    continue;
                }else {
                    String content1 = temp.getCell(1).toString();
                    if ( content1.equals("")){
                        continue;
                    }
                    Bin bin = new Bin();
                    count++;
                    Row tempRow = tempSheet.getRow(i);
                    for (int j = 0; j < Constant.TONGZHANCOL; j++) {
                        if (j == 0){
                            bin.add(String.valueOf(count));
                        }else if (j >= 3){
                            if (tempRow.getCell(j).equals(null)){
                                continue;
                            }
                            if (tempRow.getCell(j).getCellType() == Cell.CELL_TYPE_NUMERIC){
                                bin.add(String.valueOf((int) tempRow.getCell(j).getNumericCellValue()));
                            }else {
                                bin.add(tempRow.getCell(j).toString());
                            }
                        }else {
                            bin.add(tempRow.getCell(j).toString());
                        }
                    }
                    dataList.add(bin);
                }
            }
        }
        //将datalist中的数据存入data中
        data = new String[dataList.size() + 1][Constant.TONGZHANCOL];
        for (int i = 1; i < data.length; i++) {
            for (int j = 0; j < dataList.get(i - 1).size(); j++) {
                data[i][j] = dataList.get(i - 1).get(j);
            }
        }
        String[] tempArray = calData(dataList.size());
        for (int i = 0; i < Constant.TONGZHANCOL; i++) {
            data[0][i] = tempArray[i];
        }
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
            if (name.contains("统战工作")){
                fileName = name;
                break;
            }
        }
        return fileName;
    }

    private String[] calData(int number) {
        //对data中的所有数据进行统计
        int zhuanzhi = 0;
        int dangwai = 0;
        int gaocengci = 0;
        String[] tempArray = new String[Constant.TONGZHANCOL];
        tempArray[0] = "总计";
        tempArray[1] = "--";
        tempArray[2] = "--";
        for (int i = 0; i < number; i++) {
            for (int j = 3; j < Constant.TONGZHANCOL; j++) {
                if (j == 3){
                    if (isInteger(data[i + 1][j]) || isDouble(data[i + 1][j])){
                        if (isDouble(data[i + 1][j])){
                            double temp = Double.parseDouble(data[i + 1][j]);
                            zhuanzhi += (int) temp;
                        }else {
                            zhuanzhi += Integer.valueOf(data[i + 1][j]);
                        }
                    }else {
                        zhuanzhi += 0;
                    }
                }else if (j == 4){
                    if (isInteger(data[i + 1][j]) || isDouble(data[i + 1][j])){
                        if (isDouble(data[i + 1][j])){
                            double temp = Double.parseDouble(data[i + 1][j]);
                            dangwai += (int) temp;
                        }else {
                            dangwai += Integer.valueOf(data[i + 1][j]);
                        }
                    }else {
                        dangwai += 0;
                    }
                }else {
                    if (isInteger(data[i + 1][j]) || isDouble(data[i + 1][j])){
                        if (isDouble(data[i + 1][j])){
                            double temp = Double.parseDouble(data[i + 1][j]);
                            gaocengci += (int) temp;
                        }else {
                            gaocengci += Integer.valueOf(data[i + 1][j]);
                        }
                    }else {
                        gaocengci += 0;
                    }
                }
            }
        }
        tempArray[3] = String.valueOf(zhuanzhi);
        tempArray[4] = String.valueOf(dangwai);
        tempArray[5] = String.valueOf(gaocengci);
        return tempArray;
    }

    public String[][] getData() {
        return data;
    }
}
