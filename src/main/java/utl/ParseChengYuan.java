package utl;

import org.apache.poi.ss.usermodel.*;

import java.io.File;
import java.util.ArrayList;
import java.util.List;

import static java.io.File.separator;

/**
 * describe:
 *
 * @author phantom
 * @date 2019/06/25
 */
public class ParseChengYuan {
    private String year;
    private String organization;
    private String[][] data;

    private void parseOneExcel() {
        //创建存放数据的列表
        List<Bin> dataList = new ArrayList<Bin>();
        //获取存在统战工作数据表的单位
        String parentDir = new GetYears().getPath() + separator +
                this.year + separator + organization;
        String name = getTargetFileName(parentDir);
        String path = parentDir + separator + name;


        File file = new File(path);
        String[] tempArray = {"民革","民盟","民建","民进","农工","致公党","九三","台盟"};
        if (!file.exists()){
            for (int i = 0; i < 8; i++) {
                Bin bin = new Bin();
                bin.add(tempArray[i]);
                for (int j = 1; j < Constant.ChengYuan; j++) {
                    bin.add("");
                }
                dataList.add(bin);
            }
            updateData(dataList);
            return;
        }

        //获得操作的对象
        Workbook wb = GetWorkBook.getWorkBook(this.year, this.organization, name);
        //以上代码是为了获取每一个目标文件的操作对象，接下来读取每一个文件中的内容，存入文件中
        Sheet tempSheet = wb.getSheetAt(0);
        //获取统战工作表中需要统计的信息所在的行
        for (int i = 4; i < 13; i++) {
            Row temp = tempSheet.getRow(i);
            Bin bin = new Bin();
            if (i != 12){
                for (int j = 0; j < Constant.ChengYuan; j++) {
                    if (j >= 1 ){
                        if (temp.getCell(j).equals(null) || temp.getCell(j).equals("")){
                            bin.add(String.valueOf(0));
                        }else {
                            if (temp.getCell(j).getCellType() == Cell.CELL_TYPE_NUMERIC){
                                bin.add(String.valueOf((int) temp.getCell(j).getNumericCellValue()));
                            }else {
                                bin.add(temp.getCell(j).toString());
                            }
                        }
                    }else {
                        bin.add(temp.getCell(j).toString());
                    }

                }
            }else {
                for (int j = 0; j < Constant.ChengYuan; j++) {
                    if (j == 0){
                        String tempContent = "";
                        if (!temp.getCell(j).equals(null)){
                            tempContent = temp.getCell(j).toString();
                        }
                        bin.add(tempContent);
                    }else {
                        if (temp.getCell(j).getCellType() == Cell.CELL_TYPE_FORMULA){
                           bin.add(String.valueOf((int)temp.getCell(j).getNumericCellValue()));
                        }else {
                            bin.add(temp.getCell(j).toString());
                        }
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
        //初始化9个Bin使得每个Bin的List中存放的是各个表格的累积的和
        for (int i = 0; i < 9; i++) {
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
            for (int i = 4; i < 13; i++) {
                Row temp = tempSheet.getRow(i);
                if (count == 1){
                    if (i != 12){
                        for (int j = 0; j < Constant.ChengYuan; j++) {
                            String tempContent = "";
                            if (!temp.getCell(j).equals(null)){
                                tempContent = temp.getCell(j).toString();
                            }
                            dataList.get(i - 4).add(tempContent);
                        }
                    }else {
                        for (int j = 0; j < Constant.ChengYuan; j++) {
                            if (j == 0){
                                String tempContent = "";
                                if (!temp.getCell(j).equals(null)){
                                    tempContent = temp.getCell(j).toString();
                                }
                                dataList.get(i - 4).add(tempContent);
                            }else {
                                if (temp.getCell(j).getCellType() == Cell.CELL_TYPE_FORMULA){
                                    dataList.get(i - 4).add(String.valueOf(temp.getCell(j).getNumericCellValue()));
                                }else {
                                    dataList.get(i - 4).add(temp.getCell(j).toString());
                                }
                            }
                        }
                    }
                }else { //在第一个读取的lists上增加内容
                    if (i != 12){
                        for (int j = 0; j < Constant.ChengYuan; j++) {
                            if (j == 0){
                                continue;
                            }
                            if (temp.getCell(j).getCellType() == Cell.CELL_TYPE_NUMERIC){
                                double content1 = Double.parseDouble(temp.getCell(j).toString());
                                double content2 = 0;
                                try {
                                    content2 = Double.parseDouble(dataList.get(i - 4).get(j));
                                }catch (NumberFormatException e){
                                    content2 = 0;
                                }
                                dataList.get(i - 4).changeValue(j, String.valueOf((int) content1 + (int)content2));
                            }
                        }
                    }else {
                        for (int j = 0; j < Constant.ChengYuan; j++) {
                            if (j == 0){
                                continue;
                            }else {
                                if (temp.getCell(j).getCellType() == Cell.CELL_TYPE_FORMULA){
                                    double content1 = temp.getCell(j).getNumericCellValue();
                                    double content2 = Double.parseDouble(dataList.get(i - 4).get(j));
                                    dataList.get(i - 4).changeValue(j, String.valueOf((int) content1 + (int)content2));
                                }else {
                                    double content1 = 0;
                                    try{
                                        content1 = Double.parseDouble(temp.getCell(j).toString());
                                    }catch (NumberFormatException e){
                                        content1 = 0;
                                    }
                                    double content2 = 0;
                                    try{
                                        content2 = Double.parseDouble(dataList.get(i - 4).get(j));
                                    }catch (NumberFormatException e){
                                        content2 = 0;
                                    }
                                    dataList.get(i - 4).changeValue(j, String.valueOf((int) content1 + (int)content2));
                                }
                            }
                        }
                    }
                }
            }
        }
        //将datalist中的数据存入data中
        updateData(dataList);
    }

    public String[][] getData() {
        return data;
    }

    /**
     * 将列表中的数据存入data中
     * @param dataList 列表中的数据
     */
    private void updateData(List<Bin> dataList) {
        data = new String[dataList.size()][Constant.ChengYuan];
        for (int i = 0; i < data.length; i++) {
            for (int j = 0; j < Constant.ChengYuan; j++) {
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
            if (name.contains("民主党派")){
                fileName = name;
                break;
            }
        }
        return fileName;
    }

    public ParseChengYuan(String year, String organization) {
        this.year = year;
        this.organization = organization;
        if (organization.equals("所有学校")){
            parseAllExcel();
        }else {
            parseOneExcel();
        }
    }

    public static void main(String[] args) {
//        ParseChengYuan parseChengYuan = new ParseChengYuan("2019", "北京工业大学");
        ParseChengYuan parseChengYuan = new ParseChengYuan("2019", "所有学校");
        String[][] data = parseChengYuan.getData();
        for (int i = 0; i < data.length; i++) {
            for (int j = 0; j < data[0].length; j++) {
                System.out.print(data[i][j] + ", ");
            }
            System.out.println();
        }
    }

}
