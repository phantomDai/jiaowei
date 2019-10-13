package utl;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.File;
import java.util.ArrayList;
import java.util.List;

import static java.io.File.separator;
import static utl.CommonMethods.isDouble;
import static utl.CommonMethods.isInteger;

/**
 * describe:
 *
 * @author phantom
 * @date 2019/06/25
 */
public class ParseZhiLianHui {
    private String year;
    private String organization;
    private String[][] data;

    public String[][] getData() {
        return data;
    }

    public ParseZhiLianHui(String year, String organization) {
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

        File file = new File(filePath);
        if (!file.exists()){
            Bin bin = new Bin();
            for (int i = 0; i < Constant.ZhiLian; i++) {
                bin.add("");
            }
            dataList.add(bin);
            updateData(dataList);
            return;
        }

        //获得操作的对象
        Workbook wb = GetWorkBook.getWorkBook(filePath);


        //以上代码是为了获取每一个目标文件的操作对象，接下来读取每一个文件中的内容，存入文件中
        Sheet tempSheet = wb.getSheetAt(0);
        //获取统战工作表中需要统计的信息所在的行
        int numberOfRows = tempSheet.getPhysicalNumberOfRows();
        int count = 0;

        for (int i = 3; i < numberOfRows; i++) {
            Row temp = tempSheet.getRow(i);
            if (temp.getCell(0).equals(null) || temp.getCell(0).toString().equals("")){
                break;
            }
            if (temp.getCell(1).equals(null)){
                continue;
            }else {
                count++;
                Bin bin = new Bin();
                for (int j = 0; j < Constant.ZhiLian; j++) {
                    if (j == 0){
                        bin.add(String.valueOf(count));
                    }else if (j == 2 || j == 6){
                        bin.add(CommonMethods.parseDate(temp.getCell(j)));
                    }else if (j == 5 || j == 9){
                        if (temp.getCell(j).equals(null) || temp.getCell(j).equals("")){
                            bin.add(String.valueOf(0));
                        }else if (temp.getCell(j).getCellType() == Cell.CELL_TYPE_NUMERIC){
                            bin.add(String.valueOf((int) temp.getCell(j).getNumericCellValue()));
                        }else {
                            bin.add(temp.getCell(j).toString());
                        }
                    }else {
                        bin.add(temp.getCell(j).toString());
                    }
                }
                dataList.add(bin);
            }
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
                if (excelName.contains("知联会")){
                    organizations.add(item);
                    break;
                }
            }
        }
        List<Bin> dataList = new ArrayList<Bin>();
        int count = 0; //计数器
        //逐个读取包含“知联会”的单位的表格信息
        for (String organization: organizations) {
            String tempOrganizationDir = parentDir + separator + organization;

            String fileName = getTargetFileName(tempOrganizationDir);

            String filePath = tempOrganizationDir + separator + fileName;

            Workbook wb = GetWorkBook.getWorkBook(filePath);
            Sheet tempSheet = wb.getSheetAt(0);

            int numberOfRows = tempSheet.getPhysicalNumberOfRows();
            for (int i = 3; i < numberOfRows; i++) {
                Row temp = tempSheet.getRow(i);
                if (temp.getCell(0).equals(null) || temp.getCell(0).toString().equals("")){
                    break;
                }
                if (temp.getCell(1).equals(null)){
                    continue;
                }else {
                    count++;
                    Bin bin = new Bin();
                    for (int j = 0; j < Constant.ZhiLian; j++) {
                        if (j == 0){
                            bin.add(String.valueOf(count));
                        }else if (j == 2 || j == 6){
                            bin.add(CommonMethods.parseDate(temp.getCell(j)));
                        }else if (j == 5 || j == 9){
                            if (temp.getCell(j).equals(null) || temp.getCell(j).equals("")){
                                bin.add(String.valueOf(0));
                            }else if (temp.getCell(j).getCellType() == Cell.CELL_TYPE_NUMERIC){
                                bin.add(String.valueOf((int) temp.getCell(j).getNumericCellValue()));
                            }else {
                                bin.add(temp.getCell(j).toString());
                            }
                        }else {
                            bin.add(temp.getCell(j).toString());
                        }
                    }
                    dataList.add(bin);
                }
            }
        }
        //将datalist中的数据存入data中
        updateData(dataList);
    }

    /**
     * 将列表中的数据存入data中
     * @param dataList 列表中的数据
     */
    private void updateData(List<Bin> dataList) {
        data = new String[dataList.size()][Constant.ZhiLian];
        for (int i = 0; i < data.length; i++) {
            for (int j = 0; j < Constant.ZhiLian; j++) {
                data[i][j] = dataList.get(i).get(j);
            }
        }
    }


    /**
     * 给出父目录获取指定的文件名字
     * @param parentDir 父目录
     * @return 指定的文件的名字
     */
    private String getTargetFileName(String parentDir) {
        String fileName = "不存在";
        String[] tempNames = new File(parentDir).list();
        for (String name: tempNames) {
            if (name.contains("知联会")){
                fileName = name;
                break;
            }
        }
        return fileName;
    }

    public static void main(String[] args) {
//        ParseZhiLianHui parseZhiLianHui = new ParseZhiLianHui("2019", "北京工业大学");
        ParseZhiLianHui parseZhiLianHui = new ParseZhiLianHui("2019", "所有学校");
        String[][] data = parseZhiLianHui.getData();
        for (int i = 0; i < data.length; i++) {
            for (int j = 0; j < data[0].length; j++) {
                System.out.print(data[i][j] + ", ");
            }
            System.out.println();
        }
    }
}
