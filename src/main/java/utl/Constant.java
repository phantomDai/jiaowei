package utl;

/**
 * describe:
 *
 * @author phantom
 * @date 2019/06/24
 */
public class Constant {
    /**统战文档的列数*/
    public static final int TONGZHANCOL = 6;

    /**局级文档的列数*/
    public static final int JUJI = 10;

    /**高级知识分子文档的列数*/
    public static final int ZHISHI = 12;

    public static final int ChengYuan = 24;

    public static final int ZuZhiJiGou = 7;

    public static final int ShiZhiGanBu = 16;

    public static final int ZhiLian = 10;

    public static final int ZuZhiFaZhan = 15;

    /**统战表格的表头*/
    public static final String[] titleTongZhan = {"序号","高校", "是否设置单独统战部门", "专职干部人数",
            "党外代表数量", "高层次党外代表数量"};

    public static final String[] titileJuJi = {
            "序号", "姓名", "单位", "现任职务", "出生年月", "籍贯", "党派", "学历", "职称", "任职时间"
    };

    public static final String[] titileZhiShi = {
            "年龄段", "党外高级知识分子总数", "党外正高（男）", "党外正高（女）", "党外副高（男）",
            "党外副高（女）", "高级知识分子总数", "高级知识分子（男）", "高级知识分子（女）",
            "教职工总数", "教职工（男）", "教职工（女）"
    };

    public static final String[] titleChengYuan = {
            "党派名称", "成员总数", "成员（男）", "成员（女）", "职称（高级）", "职称（中级）",
            "职称（初级）", "职称（无）", "年龄（29岁以下）", "年龄（30-39岁）", "年龄（40-49岁）",
            "年龄（50-59岁）", "年龄（60岁以上）", "离退休", "交叉党员（总数）", "交叉党员（在职数）",
            "党派市委委员以上", "党派中央委员以上", "担任处级职务", "担任局级以上职务", "变化情况（调入数）",
            "变化情况（调出数）", "变化情况（退出数）", "变化情况（死亡数）"
    };

    public static final String[] titleZuZhiJiGou = {
            "党派名称", "委员会数", "总支数", "支部数", "小组数", "总人数", "联合组织数"
    };

    public static final String[] titleShiZhi = {
            "类别", "党外实职干部（总数）", "党外实职干部（男）", "党外实职干部（女）",
            "党外实职干部（高级职称）", "党外实职干部（中级职称）", "党外实职干部（初级职称）",
            "党外实职干部（无职称）", "党外实职干部（29岁以下）", "党外实职干部（30-39岁）",
            "党外实职干部（40-49岁）", "党外实职干部（50-59岁）", "党外实职干部（60岁以上）",
            "党外实职干部（少数民族）", "党外实职干部（民主党派）", "实职干部总数"
    };

    public static final String[] titleZhiLian = {
            "序号", "学校", "知识分子联谊会成立时间", "知识分子联谊会会长姓名", "知识分子联谊会会长职务",
            "知识分子联谊会成员人数", "归国人员联谊会成立时间", "归国人员联谊会会长姓名",
            "归国人员联谊会会长职务", "归国人员联谊会成员人数"
    };

    public static final String[] titleZuZhiFaZhan = {
            "党派名称", "新建组织（委员会数）", "新建组织（总支数）", "新建组织（支部数）",
            "新建组织（小组数）", "发展成员数量", "职称（高级）", "职称（中级）", "职称（初级）",
            "职称（无）", "年龄（29岁以下）", "年龄（30-39岁）", "年龄（40-49岁）",
            "年龄（50-59岁）", "年龄（60岁以上）"
    };

}
