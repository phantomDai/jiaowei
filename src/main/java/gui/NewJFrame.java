/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package gui;

import utl.*;

import javax.swing.*;
import javax.swing.table.DefaultTableModel;
import javax.swing.table.TableModel;

/**
 *
 * @author daihe
 */
public class NewJFrame extends javax.swing.JFrame {
    /**
     * @param args the command line arguments
     */
    public static void main(String args[]) {
        /* Set the Nimbus look and feel */
        //<editor-fold defaultstate="collapsed" desc=" Look and feel setting code (optional) ">
        /* If Nimbus (introduced in Java SE 6) is not available, stay with the default look and feel.
         * For details see http://download.oracle.com/javase/tutorial/uiswing/lookandfeel/plaf.html
         */
        try {
            for (javax.swing.UIManager.LookAndFeelInfo info : javax.swing.UIManager.getInstalledLookAndFeels()) {
                if ("Nimbus".equals(info.getName())) {
                    javax.swing.UIManager.setLookAndFeel(info.getClassName());
                    break;
                }
            }
        } catch (ClassNotFoundException ex) {
            java.util.logging.Logger.getLogger(NewJFrame.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (InstantiationException ex) {
            java.util.logging.Logger.getLogger(NewJFrame.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (IllegalAccessException ex) {
            java.util.logging.Logger.getLogger(NewJFrame.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (javax.swing.UnsupportedLookAndFeelException ex) {
            java.util.logging.Logger.getLogger(NewJFrame.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        }
        //</editor-fold>

        /* Create and display the form */
        java.awt.EventQueue.invokeLater(new Runnable() {
            public void run() {
                new NewJFrame().setVisible(true);
            }
        });
    }

    // Variables declaration - do not modify
    private javax.swing.JPanel buZhoujPanel;
    private javax.swing.JPanel chengyuanjPanel6;
    private javax.swing.JTable chengyuanjTable1;
    private javax.swing.JComboBox<String> danweijComboBox1;
    private javax.swing.JLabel danweijLabel6;
    private javax.swing.JButton jButton1;
    private javax.swing.JButton jButton2;
    private javax.swing.JLabel jLabel1;
    private javax.swing.JLabel jLabel2;
    private javax.swing.JLabel jLabel3;
    private javax.swing.JLabel jLabel4;
    private javax.swing.JLabel jLabel5;
    private javax.swing.JPanel jPanel1;
    private javax.swing.JScrollPane jScrollPane1;
    private javax.swing.JScrollPane jScrollPane2;
    private javax.swing.JScrollPane jScrollPane3;
    private javax.swing.JScrollPane jScrollPane4;
    private javax.swing.JScrollPane jScrollPane5;
    private javax.swing.JScrollPane jScrollPane6;
    private javax.swing.JScrollPane jScrollPane7;
    private javax.swing.JScrollPane jScrollPane8;
    private javax.swing.JTabbedPane jTabbedPane1;
    private javax.swing.JPanel jujijPanel4;
    private javax.swing.JTable jujijTable1;
    private javax.swing.JComboBox<String> nianduComboBox1;
    private javax.swing.JLabel niandujLabel;
    private javax.swing.JPanel shizhiganbujPanel9;
    private javax.swing.JTable shizhijTable1;
    private javax.swing.JButton shuaxinjButton1;
    private javax.swing.JPanel tongzhanjPanel;
    private javax.swing.JTable tongzhanjTable1;
    private javax.swing.JPanel zhilianhuijPanel2;
    private javax.swing.JTable zhilianjTable1;
    private javax.swing.JPanel zhishijPanel5;
    private javax.swing.JTable zhishijTable1;
    private javax.swing.JTable zuhifazhanTable1;
    private javax.swing.JPanel zuzhifazhanjPanel7;
    private javax.swing.JPanel zuzhijigoujPanel8;
    private javax.swing.JTable zuzhijigoujTable1;
    // End of variables declaration

    /**
     * Creates new form NewJFrame
     */
    public NewJFrame() {
        initComponents();
    }

    /**
     * This method is called from within the constructor to initialize the form.
     * WARNING: Do NOT modify this code. The content of this method is always
     * regenerated by the Form Editor.
     */
    @SuppressWarnings("unchecked")
    // <editor-fold defaultstate="collapsed" desc="Generated Code">
    private void initComponents() {

        buZhoujPanel = new javax.swing.JPanel();
        jLabel1 = new javax.swing.JLabel();
        jLabel2 = new javax.swing.JLabel();
        jLabel3 = new javax.swing.JLabel();
        jLabel4 = new javax.swing.JLabel();
        jLabel5 = new javax.swing.JLabel();
        jPanel1 = new javax.swing.JPanel();
        niandujLabel = new javax.swing.JLabel();
        nianduComboBox1 = new javax.swing.JComboBox<>();
        danweijLabel6 = new javax.swing.JLabel();
        danweijComboBox1 = new javax.swing.JComboBox<>();
        shuaxinjButton1 = new javax.swing.JButton();
        jButton1 = new javax.swing.JButton();
        jButton2 = new javax.swing.JButton();
        jTabbedPane1 = new javax.swing.JTabbedPane();
        tongzhanjPanel = new javax.swing.JPanel();
        jScrollPane1 = new javax.swing.JScrollPane();
        tongzhanjTable1 = new javax.swing.JTable();
        jujijPanel4 = new javax.swing.JPanel();
        jScrollPane2 = new javax.swing.JScrollPane();
        jujijTable1 = new javax.swing.JTable();
        zhishijPanel5 = new javax.swing.JPanel();
        jScrollPane3 = new javax.swing.JScrollPane();
        zhishijTable1 = new javax.swing.JTable();
        chengyuanjPanel6 = new javax.swing.JPanel();
        jScrollPane4 = new javax.swing.JScrollPane();
        chengyuanjTable1 = new javax.swing.JTable();
        zuzhijigoujPanel8 = new javax.swing.JPanel();
        jScrollPane6 = new javax.swing.JScrollPane();
        zuzhijigoujTable1 = new javax.swing.JTable();
        shizhiganbujPanel9 = new javax.swing.JPanel();
        jScrollPane7 = new javax.swing.JScrollPane();
        shizhijTable1 = new javax.swing.JTable();
        zhilianhuijPanel2 = new javax.swing.JPanel();
        jScrollPane8 = new javax.swing.JScrollPane();
        zhilianjTable1 = new javax.swing.JTable();
        zuzhifazhanjPanel7 = new javax.swing.JPanel();
        jScrollPane5 = new javax.swing.JScrollPane();
        zuhifazhanTable1 = new javax.swing.JTable();

        /**一下部分是对GUI窗体作一些简单的修改***************************************/
        setDefaultCloseOperation(javax.swing.WindowConstants.EXIT_ON_CLOSE);
        setResizable(false);
        ImageIcon logo = new ImageIcon("src/image/logo.png");
        this.setIconImage(logo.getImage());
        setTitle("北京市高校党外人员信息统计");
        /**********************************************************************/

        buZhoujPanel.setBorder(javax.swing.BorderFactory.createTitledBorder("操作步骤"));

        jLabel1.setText("步骤2：选择统计信息的年份");

        jLabel2.setText("步骤1：点击“刷新”按钮");

        jLabel3.setText("步骤3：选择统计信息的单位");

        jLabel4.setText("步骤4：点击“查询”按钮");

        jLabel5.setText("步骤5：点击不同的选项卡查询信息");

        javax.swing.GroupLayout buZhoujPanelLayout = new javax.swing.GroupLayout(buZhoujPanel);
        buZhoujPanel.setLayout(buZhoujPanelLayout);
        buZhoujPanelLayout.setHorizontalGroup(
                buZhoujPanelLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                        .addGroup(buZhoujPanelLayout.createSequentialGroup()
                                .addGroup(buZhoujPanelLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                                        .addComponent(jLabel1)
                                        .addComponent(jLabel2)
                                        .addComponent(jLabel3)
                                        .addComponent(jLabel4)
                                        .addComponent(jLabel5))
                                .addGap(0, 0, Short.MAX_VALUE))
        );
        buZhoujPanelLayout.setVerticalGroup(
                buZhoujPanelLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                        .addGroup(buZhoujPanelLayout.createSequentialGroup()
                                .addContainerGap()
                                .addComponent(jLabel2)
                                .addGap(30, 30, 30)
                                .addComponent(jLabel1)
                                .addGap(26, 26, 26)
                                .addComponent(jLabel3)
                                .addGap(28, 28, 28)
                                .addComponent(jLabel4)
                                .addGap(31, 31, 31)
                                .addComponent(jLabel5)
                                .addContainerGap(javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE))
        );

        jPanel1.setBorder(new javax.swing.border.SoftBevelBorder(javax.swing.border.BevelBorder.RAISED));

        niandujLabel.setFont(new java.awt.Font("宋体", 0, 20)); // NOI18N
        niandujLabel.setText("年度：");

        nianduComboBox1.setFont(new java.awt.Font("宋体", 0, 20)); // NOI18N
        nianduComboBox1.setModel(new javax.swing.DefaultComboBoxModel<>(new String[] { "Item 1", "Item 2", "Item 3", "Item 4" }));

        danweijLabel6.setFont(new java.awt.Font("宋体", 0, 20)); // NOI18N
        danweijLabel6.setText("单位：");

        danweijComboBox1.setFont(new java.awt.Font("宋体", 0, 20)); // NOI18N
        danweijComboBox1.setModel(new javax.swing.DefaultComboBoxModel<>(new String[] { "Item 1", "Item 2", "Item 3", "Item 4" }));

        shuaxinjButton1.setFont(new java.awt.Font("宋体", 0, 20)); // NOI18N
        shuaxinjButton1.setText("刷新");
        shuaxinjButton1.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                shuaxinjButton1ActionPerformed(evt);
            }
        });

        jButton1.setFont(new java.awt.Font("宋体", 0, 20)); // NOI18N
        jButton1.setText("查询");
        jButton1.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                chaXunjButton1ActionPerformed(evt);
            }
        });


        jButton2.setFont(new java.awt.Font("宋体", 0, 20)); // NOI18N
        jButton2.setText("导出");
        jButton2.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jButton2ActionPerformed(evt);
            }
        });

        javax.swing.GroupLayout jPanel1Layout = new javax.swing.GroupLayout(jPanel1);
        jPanel1.setLayout(jPanel1Layout);
        jPanel1Layout.setHorizontalGroup(
                jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                        .addGroup(jPanel1Layout.createSequentialGroup()
                                .addContainerGap()
                                .addComponent(niandujLabel)
                                .addGap(50, 50, 50)
                                .addComponent(nianduComboBox1, javax.swing.GroupLayout.PREFERRED_SIZE, 110, javax.swing.GroupLayout.PREFERRED_SIZE)
                                .addGap(109, 109, 109)
                                .addComponent(danweijLabel6)
                                .addGap(54, 54, 54)
                                .addComponent(danweijComboBox1, javax.swing.GroupLayout.PREFERRED_SIZE, 140, javax.swing.GroupLayout.PREFERRED_SIZE)
                                .addGap(67, 67, 67)
                                .addComponent(shuaxinjButton1)
                                .addGap(39, 39, 39)
                                .addComponent(jButton1)
                                .addGap(32, 32, 32)
                                .addComponent(jButton2)
                                .addContainerGap(javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE))
        );
        jPanel1Layout.setVerticalGroup(
                jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                        .addGroup(jPanel1Layout.createSequentialGroup()
                                .addGap(29, 29, 29)
                                .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                                        .addComponent(niandujLabel)
                                        .addComponent(nianduComboBox1, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                                        .addComponent(danweijLabel6)
                                        .addComponent(danweijComboBox1, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                                        .addComponent(shuaxinjButton1)
                                        .addComponent(jButton1)
                                        .addComponent(jButton2))
                                .addContainerGap(39, Short.MAX_VALUE))
        );

        jTabbedPane1.setBorder(javax.swing.BorderFactory.createEtchedBorder());
        jTabbedPane1.setFont(new java.awt.Font("宋体", 0, 20)); // NOI18N

        tongzhanjTable1.setFont(new java.awt.Font("宋体", 0, 20)); // NOI18N
        tongzhanjTable1.setModel(new javax.swing.table.DefaultTableModel(
                new Object [][] {
                        {null, null, null, null, null, null},
                        {null, null, null, null, null, null}
                },
                new String [] {
                        "序号", "高校", "是否设置单独统战部门", "专职干部数", "党外代表数量", "高层次党外代表数量"
                }
        ));
        tongzhanjTable1.setAutoResizeMode(javax.swing.JTable.AUTO_RESIZE_OFF);
        tongzhanjTable1.setRowHeight(30);
        for (int i = 0; i < Constant.TONGZHANCOL; i++) {
            tongzhanjTable1.getColumnModel().getColumn(i).setPreferredWidth(155);
        }
        jScrollPane1.setViewportView(tongzhanjTable1);

        javax.swing.GroupLayout tongzhanjPanelLayout = new javax.swing.GroupLayout(tongzhanjPanel);
        tongzhanjPanel.setLayout(tongzhanjPanelLayout);
        tongzhanjPanelLayout.setHorizontalGroup(
                tongzhanjPanelLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                        .addComponent(jScrollPane1, javax.swing.GroupLayout.DEFAULT_SIZE, 937, Short.MAX_VALUE)
        );
        tongzhanjPanelLayout.setVerticalGroup(
                tongzhanjPanelLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                        .addComponent(jScrollPane1, javax.swing.GroupLayout.DEFAULT_SIZE, 642, Short.MAX_VALUE)
        );

        jTabbedPane1.addTab("统战工作", tongzhanjPanel);

        jujijTable1.setFont(new java.awt.Font("宋体", 0, 20)); // NOI18N
        jujijTable1.setModel(new javax.swing.table.DefaultTableModel(
                new Object [][] {
                        {null, null, null, null, null, null, null, null, null, null},
                        {null, null, null, null, null, null, null, null, null, null},
                        {null, null, null, null, null, null, null, null, null, null},
                        {null, null, null, null, null, null, null, null, null, null}
                },
                new String [] {
                        "序号", "姓名", "单位", "现任职务", "出生年月", "籍贯", "党派", "学历", "职称", "任职时间"
                }
        ));
        jujijTable1.setAutoResizeMode(javax.swing.JTable.AUTO_RESIZE_OFF);
        jujijTable1.setRowHeight(30);
        for (int i = 0; i < Constant.JUJI; i++) {
            jujijTable1.getColumnModel().getColumn(i).setPreferredWidth(93);
        }
        jScrollPane2.setViewportView(jujijTable1);

        javax.swing.GroupLayout jujijPanel4Layout = new javax.swing.GroupLayout(jujijPanel4);
        jujijPanel4.setLayout(jujijPanel4Layout);
        jujijPanel4Layout.setHorizontalGroup(
                jujijPanel4Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                        .addComponent(jScrollPane2, javax.swing.GroupLayout.DEFAULT_SIZE, 937, Short.MAX_VALUE)
        );
        jujijPanel4Layout.setVerticalGroup(
                jujijPanel4Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                        .addComponent(jScrollPane2, javax.swing.GroupLayout.DEFAULT_SIZE, 642, Short.MAX_VALUE)
        );

        jTabbedPane1.addTab("局级干部", jujijPanel4);

        zhishijTable1.setFont(new java.awt.Font("宋体", 0, 20)); // NOI18N
        zhishijTable1.setModel(new javax.swing.table.DefaultTableModel(
                new Object [][] {
                        {null, null, null, null, null, null, null, null, null, null, null, null},
                        {null, null, null, null, null, null, null, null, null, null, null, null},
                        {null, null, null, null, null, null, null, null, null, null, null, null},
                        {null, null, null, null, null, null, null, null, null, null, null, null},
                        {null, null, null, null, null, null, null, null, null, null, null, null},
                        {null, null, null, null, null, null, null, null, null, null, null, null}
                },
                new String [] {
                        "年龄段", "党外高级知识分子总数", "党外正高（男）", "党外正高（女）", "党外副高（男）", "党外副高（女）", "高级知识分子总数", "高级知识分子（男）", "高级知识分子（女）", "教职工总数", "教职工（男）", "教职工（女）"
                }
        ));
        zhishijTable1.setAutoResizeMode(javax.swing.JTable.AUTO_RESIZE_OFF);
        zhishijTable1.setRowHeight(30);
        for (int i = 0; i < Constant.ZHISHI; i++) {
            if (i == 0){
                zhishijTable1.getColumnModel().getColumn(i).setPreferredWidth(100);
            }else if (i == 1){
                zhishijTable1.getColumnModel().getColumn(i).setPreferredWidth(140);
            }else {
                zhishijTable1.getColumnModel().getColumn(i).setPreferredWidth(120);
            }

        }
        jScrollPane3.setViewportView(zhishijTable1);

        javax.swing.GroupLayout zhishijPanel5Layout = new javax.swing.GroupLayout(zhishijPanel5);
        zhishijPanel5.setLayout(zhishijPanel5Layout);
        zhishijPanel5Layout.setHorizontalGroup(
                zhishijPanel5Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                        .addComponent(jScrollPane3, javax.swing.GroupLayout.DEFAULT_SIZE, 937, Short.MAX_VALUE)
        );
        zhishijPanel5Layout.setVerticalGroup(
                zhishijPanel5Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                        .addGroup(zhishijPanel5Layout.createSequentialGroup()
                                .addComponent(jScrollPane3, javax.swing.GroupLayout.PREFERRED_SIZE, 225, javax.swing.GroupLayout.PREFERRED_SIZE)
                                .addGap(0, 417, Short.MAX_VALUE))
        );

        jTabbedPane1.addTab("高级知识分子", zhishijPanel5);

        chengyuanjTable1.setFont(new java.awt.Font("宋体", 0, 20)); // NOI18N
        chengyuanjTable1.setModel(new javax.swing.table.DefaultTableModel(
                new Object [][] {
                        {"民革", null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null},
                        {"民盟", null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null},
                        {"民建", null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null},
                        {"民进", null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null},
                        {"农工", null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null},
                        {"致公党", null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null},
                        {"九三", null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null},
                        {"台盟", null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null},
                        {null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null}
                },
                new String [] {
                        "党派名称", "成员总数", "成员（男）", "成员（女）", "职称（高级）", "职称（中级）", "职称（初级）", "职称（无）", "年龄（29岁以下）", "年龄（30-39岁）", "年龄（40-49岁）", "年龄（50-59岁）", "年龄（60岁以上）", "离退休", "交叉党员（总数）", "交叉党员（在职数）", "党派市委委员以上", "党派中央委员以上", "担任处级职务", "担任局级以上职务", "变化情况（调入数）", "变化情况（调出数）", "变化情况（退出数）", "变化情况（死亡数）"
                }
        ));
        chengyuanjTable1.setAutoResizeMode(javax.swing.JTable.AUTO_RESIZE_OFF);
        chengyuanjTable1.setRowHeight(30);
        for (int i = 0; i < Constant.ChengYuan; i++) {
            chengyuanjTable1.getColumnModel().getColumn(i).setPreferredWidth(120);
        }
        jScrollPane4.setViewportView(chengyuanjTable1);

        javax.swing.GroupLayout chengyuanjPanel6Layout = new javax.swing.GroupLayout(chengyuanjPanel6);
        chengyuanjPanel6.setLayout(chengyuanjPanel6Layout);
        chengyuanjPanel6Layout.setHorizontalGroup(
                chengyuanjPanel6Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                        .addComponent(jScrollPane4, javax.swing.GroupLayout.DEFAULT_SIZE, 937, Short.MAX_VALUE)
        );
        chengyuanjPanel6Layout.setVerticalGroup(
                chengyuanjPanel6Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                        .addGroup(chengyuanjPanel6Layout.createSequentialGroup()
                                .addComponent(jScrollPane4, javax.swing.GroupLayout.PREFERRED_SIZE, 336, javax.swing.GroupLayout.PREFERRED_SIZE)
                                .addGap(0, 306, Short.MAX_VALUE))
        );

        jTabbedPane1.addTab("党派成员", chengyuanjPanel6);

        zuzhijigoujTable1.setFont(new java.awt.Font("宋体", 0, 20)); // NOI18N
        zuzhijigoujTable1.setModel(new javax.swing.table.DefaultTableModel(
                new Object [][] {
                        {null, null, null, null, null, null, null},
                        {null, null, null, null, null, null, null},
                        {null, null, null, null, null, null, null},
                        {null, null, null, null, null, null, null},
                        {null, null, null, null, null, null, null},
                        {null, null, null, null, null, null, null},
                        {null, null, null, null, null, null, null},
                        {null, null, null, null, null, null, null},
                        {null, null, null, null, null, null, null}
                },
                new String [] {
                        "党派名称", "委员会数", "总支数", "支部数", "小组数", "总人数", "联合组织数"
                }
        ));
        zuzhijigoujTable1.setAutoResizeMode(javax.swing.JTable.AUTO_RESIZE_OFF);
        zuzhijigoujTable1.setRowHeight(30);
        for (int i = 0; i < Constant.ZuZhiJiGou; i++) {
            zuzhijigoujTable1.getColumnModel().getColumn(i).setPreferredWidth(132);
        }
        jScrollPane6.setViewportView(zuzhijigoujTable1);

        javax.swing.GroupLayout zuzhijigoujPanel8Layout = new javax.swing.GroupLayout(zuzhijigoujPanel8);
        zuzhijigoujPanel8.setLayout(zuzhijigoujPanel8Layout);
        zuzhijigoujPanel8Layout.setHorizontalGroup(
                zuzhijigoujPanel8Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                        .addComponent(jScrollPane6, javax.swing.GroupLayout.DEFAULT_SIZE, 937, Short.MAX_VALUE)
        );
        zuzhijigoujPanel8Layout.setVerticalGroup(
                zuzhijigoujPanel8Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                        .addGroup(zuzhijigoujPanel8Layout.createSequentialGroup()
                                .addComponent(jScrollPane6, javax.swing.GroupLayout.PREFERRED_SIZE, 314, javax.swing.GroupLayout.PREFERRED_SIZE)
                                .addGap(0, 328, Short.MAX_VALUE))
        );

        jTabbedPane1.addTab("组织机构", zuzhijigoujPanel8);

        shizhijTable1.setFont(new java.awt.Font("宋体", 0, 20)); // NOI18N
        shizhijTable1.setModel(new javax.swing.table.DefaultTableModel(
                new Object [][] {
                        {null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null},
                        {null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null},
                        {null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null}
                },
                new String [] {
                        "类别", "党外实职干部（总数）", "党外实职干部（男）", "党外实职干部（女）", "党外实职干部（高级职称）", "党外实职干部（中级职称）", "党外实职干部（初级职称）", "党外实职干部（无职称）", "党外实职干部（29岁以下）", "党外实职干部（30-39岁）", "党外实职干部（40-49岁）", "党外实职干部（50-59岁）", "党外实职干部（60岁以上）", "党外实职干部（少数民族）", "党外实职干部（民主党派）", "实职干部总数"
                }
        ));
        shizhijTable1.setAutoResizeMode(javax.swing.JTable.AUTO_RESIZE_OFF);
        shizhijTable1.setRowHeight(30);
        for (int i = 0; i < Constant.ShiZhiGanBu; i++) {
            if (i <= 3){
                shizhijTable1.getColumnModel().getColumn(i).setPreferredWidth(132);
            }else {
                shizhijTable1.getColumnModel().getColumn(i).setPreferredWidth(159);
            }
        }
        jScrollPane7.setViewportView(shizhijTable1);

        javax.swing.GroupLayout shizhiganbujPanel9Layout = new javax.swing.GroupLayout(shizhiganbujPanel9);
        shizhiganbujPanel9.setLayout(shizhiganbujPanel9Layout);
        shizhiganbujPanel9Layout.setHorizontalGroup(
                shizhiganbujPanel9Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                        .addComponent(jScrollPane7, javax.swing.GroupLayout.DEFAULT_SIZE, 937, Short.MAX_VALUE)
        );
        shizhiganbujPanel9Layout.setVerticalGroup(
                shizhiganbujPanel9Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                        .addGroup(shizhiganbujPanel9Layout.createSequentialGroup()
                                .addComponent(jScrollPane7, javax.swing.GroupLayout.PREFERRED_SIZE, 160, javax.swing.GroupLayout.PREFERRED_SIZE)
                                .addGap(0, 482, Short.MAX_VALUE))
        );

        jTabbedPane1.addTab("实职干部", shizhiganbujPanel9);

        zhilianjTable1.setFont(new java.awt.Font("宋体", 0, 20)); // NOI18N
        zhilianjTable1.setModel(new javax.swing.table.DefaultTableModel(
                new Object [][] {
                        {null, null, null, null, null, null, null, null, null, null},
                        {null, null, null, null, null, null, null, null, null, null}
                },
                new String [] {
                        "序号", "学校", "知识分子联谊会成立时间", "知识分子联谊会会长姓名", "知识分子联谊会会长职务", "知识分子联谊会成员人数", "归国人员联谊会成立时间", "归国人员联谊会会长姓名", "归国人员联谊会会长职务", "归国人员联谊会成员人数"
                }
        ));
        zhilianjTable1.setAutoResizeMode(javax.swing.JTable.AUTO_RESIZE_OFF);
        zhilianjTable1.setRowHeight(30);
        for (int i = 0; i < Constant.ZhiLian; i++) {
            zhilianjTable1.getColumnModel().getColumn(i).setPreferredWidth(145);
        }
        jScrollPane8.setViewportView(zhilianjTable1);

        javax.swing.GroupLayout zhilianhuijPanel2Layout = new javax.swing.GroupLayout(zhilianhuijPanel2);
        zhilianhuijPanel2.setLayout(zhilianhuijPanel2Layout);
        zhilianhuijPanel2Layout.setHorizontalGroup(
                zhilianhuijPanel2Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                        .addComponent(jScrollPane8, javax.swing.GroupLayout.DEFAULT_SIZE, 937, Short.MAX_VALUE)
        );
        zhilianhuijPanel2Layout.setVerticalGroup(
                zhilianhuijPanel2Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                        .addComponent(jScrollPane8, javax.swing.GroupLayout.DEFAULT_SIZE, 642, Short.MAX_VALUE)
        );

        jTabbedPane1.addTab("知联会", zhilianhuijPanel2);

        zuhifazhanTable1.setFont(new java.awt.Font("宋体", 0, 20)); // NOI18N
        zuhifazhanTable1.setModel(new javax.swing.table.DefaultTableModel(
                new Object [][] {
                        {null, null, null, null, null, null, null, null, null, null, null, null, null, null, null},
                        {null, null, null, null, null, null, null, null, null, null, null, null, null, null, null},
                        {null, null, null, null, null, null, null, null, null, null, null, null, null, null, null},
                        {null, null, null, null, null, null, null, null, null, null, null, null, null, null, null},
                        {null, null, null, null, null, null, null, null, null, null, null, null, null, null, null},
                        {null, null, null, null, null, null, null, null, null, null, null, null, null, null, null},
                        {null, null, null, null, null, null, null, null, null, null, null, null, null, null, null},
                        {null, null, null, null, null, null, null, null, null, null, null, null, null, null, null},
                        {null, null, null, null, null, null, null, null, null, null, null, null, null, null, null}
                },
                new String [] {
                        "党派名称", "新建组织（委员会数）", "新建组织（总支数）", "新建组织（支部数）", "新建组织（小组数）", "发展成员数量", "职称（高级）", "职称（中级）", "职称（初级）", "职称（无）", "年龄（29岁以下）", "年龄（30-39岁）", "年龄（40-49岁）", "年龄（50-59岁）", "年龄（60岁以上）"
                }
        ));
        zuhifazhanTable1.setAutoResizeMode(javax.swing.JTable.AUTO_RESIZE_OFF);
        zuhifazhanTable1.setRowHeight(30);
        for (int i = 0; i < Constant.ZuZhiFaZhan; i++) {
            zuhifazhanTable1.getColumnModel().getColumn(i).setPreferredWidth(135);
        }
        jScrollPane5.setViewportView(zuhifazhanTable1);

        javax.swing.GroupLayout zuzhifazhanjPanel7Layout = new javax.swing.GroupLayout(zuzhifazhanjPanel7);
        zuzhifazhanjPanel7.setLayout(zuzhifazhanjPanel7Layout);
        zuzhifazhanjPanel7Layout.setHorizontalGroup(
                zuzhifazhanjPanel7Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                        .addComponent(jScrollPane5, javax.swing.GroupLayout.DEFAULT_SIZE, 937, Short.MAX_VALUE)
        );
        zuzhifazhanjPanel7Layout.setVerticalGroup(
                zuzhifazhanjPanel7Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                        .addGroup(zuzhifazhanjPanel7Layout.createSequentialGroup()
                                .addComponent(jScrollPane5, javax.swing.GroupLayout.PREFERRED_SIZE, 337, javax.swing.GroupLayout.PREFERRED_SIZE)
                                .addGap(0, 305, Short.MAX_VALUE))
        );

        jTabbedPane1.addTab("组织发展", zuzhifazhanjPanel7);

        javax.swing.GroupLayout layout = new javax.swing.GroupLayout(getContentPane());
        getContentPane().setLayout(layout);
        layout.setHorizontalGroup(
                layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                        .addGroup(layout.createSequentialGroup()
                                .addComponent(buZhoujPanel, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                                .addGap(1, 1, 1)
                                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                                        .addComponent(jPanel1, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                                        .addComponent(jTabbedPane1)))
        );
        layout.setVerticalGroup(
                layout.createParallelGroup(javax.swing.GroupLayout.Alignment.TRAILING)
                        .addComponent(buZhoujPanel, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                        .addGroup(layout.createSequentialGroup()
                                .addComponent(jPanel1, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                                .addComponent(jTabbedPane1, javax.swing.GroupLayout.PREFERRED_SIZE, 684, javax.swing.GroupLayout.PREFERRED_SIZE))
        );

        pack();
    }// </editor-fold>

    private void shuaxinjButton1ActionPerformed(java.awt.event.ActionEvent evt) {
        // TODO add your handling code here:
        GetYears getYears = new GetYears();
        GetOrganization go = new GetOrganization();
        nianduComboBox1.setModel(new DefaultComboBoxModel<String>(getYears.getYears()));
        danweijComboBox1.setModel(new DefaultComboBoxModel<>(go.getOrganizations()));
    }

    private void chaXunjButton1ActionPerformed(java.awt.event.ActionEvent evt){
        /**获得选中的年份与学校*/
        String selectedYear = (String) nianduComboBox1.getSelectedItem();
        String selectedOrganization = (String) danweijComboBox1.getSelectedItem();

        /**获得解析各个文件的对象*/
        ParseTongZhan parseTongZhan = new ParseTongZhan(selectedYear,selectedOrganization);
        ParseJuji parseJuji = new ParseJuji(selectedYear, selectedOrganization);
        ParseZhiShi parseZhiShi = new ParseZhiShi(selectedYear, selectedOrganization);
        ParseChengYuan parseChengYuan = new ParseChengYuan(selectedYear, selectedOrganization);
        ParseZuZhiJiGou parseZuZhiJiGou = new ParseZuZhiJiGou(selectedYear, selectedOrganization);
        ParseShiZhiGanBu parseShiZhiGanBu = new ParseShiZhiGanBu(selectedYear, selectedOrganization);
        ParseZhiLianHui parseZhiLianHui = new ParseZhiLianHui(selectedYear, selectedOrganization);
        ParseZuZhiFaZhan parseZuZhiFaZhan = new ParseZuZhiFaZhan(selectedYear, selectedOrganization);

        /**获得解析后的数据并存入到model中*/
        tongzhanjTable1.setModel(new DefaultTableModel(parseTongZhan.getData(),
                Constant.titleTongZhan));
        setWidth("tongzhanjTable1");
        jujijTable1.setModel(new DefaultTableModel(parseJuji.getData(),
                Constant.titileJuJi));
        setWidth("jujijTable1");
        zhishijTable1.setModel(new DefaultTableModel(parseZhiShi.getData(),
                Constant.titileZhiShi));
        setWidth("zhishijTable1");
        chengyuanjTable1.setModel(new DefaultTableModel(parseChengYuan.getData(),
                Constant.titleChengYuan));
        setWidth("chengyuanjTable1");
        zuzhijigoujTable1.setModel(new DefaultTableModel(parseZuZhiJiGou.getData(),
                Constant.titleZuZhiJiGou));
        setWidth("zuzhijigoujTable1");
        shizhijTable1.setModel(new DefaultTableModel(parseShiZhiGanBu.getData(),
                Constant.titleShiZhi));
        setWidth("shizhijTable1");
        zhilianjTable1.setModel(new DefaultTableModel(parseZhiLianHui.getData(),
                Constant.titleZhiLian));
        setWidth("zhilianjTable1");
        zuhifazhanTable1.setModel(new DefaultTableModel(parseZuZhiFaZhan.getData(),
                Constant.titleZuZhiFaZhan));
        setWidth("zuhifazhanTable1");
        JOptionPane.showMessageDialog(null,"数据解析完毕！");
    }

    /**
     * 导出按钮对应的事件
     * @param evt
     */
    private void jButton2ActionPerformed(java.awt.event.ActionEvent evt) {
        // TODO add your handling code here:
        WriteTableToExcel writeTableToExcel = new WriteTableToExcel();
        writeTableToExcel.write("北京高校统战工作基本情况", getDataValue(tongzhanjTable1.getModel()));
        writeTableToExcel.write("党外局级以上领导干部（实职）情况", getDataValue(jujijTable1.getModel()));
        writeTableToExcel.write("高级知识分子状况", getDataValue(zhishijTable1.getModel()));
        writeTableToExcel.write("民主党派成员状况", getDataValue(chengyuanjTable1.getModel()));
        writeTableToExcel.write("民主党派组织发展状况", getDataValue(zuzhijigoujTable1.getModel()));
        writeTableToExcel.write("实职干部安排状况", getDataValue(shizhijTable1.getModel()));
        writeTableToExcel.write("知联会、留联会统计表（教工委）", getDataValue(zhilianjTable1.getModel()));
        writeTableToExcel.write("民主党派组织机构状况", getDataValue(zuhifazhanTable1.getModel()));
        JOptionPane.showMessageDialog(null,"数据导出完毕！");
    }


    private String[][] getDataValue(TableModel tm){
        String[][] data = new String[tm.getRowCount()][tm.getColumnCount()];
        for (int i = 0; i < tm.getRowCount(); i++) {
            for (int j = 0; j < tm.getColumnCount(); j++) {
                data[i][j] = (String) tm.getValueAt(i,j);
            }
        }
        return data;
    }

    private void setWidth(String tableName){
        if (tableName.equals("tongzhanjTable1")){
            for (int i = 0; i < Constant.TONGZHANCOL; i++) {
                tongzhanjTable1.getColumnModel().getColumn(i).setPreferredWidth(155);
            }
        }else if (tableName.equals("jujijTable1")){
            for (int i = 0; i < Constant.JUJI; i++) {
                jujijTable1.getColumnModel().getColumn(i).setPreferredWidth(93);
            }
        }else if (tableName.equals("zhishijTable1")){
            for (int i = 0; i < Constant.ZHISHI; i++) {
                if (i == 0){
                    zhishijTable1.getColumnModel().getColumn(i).setPreferredWidth(100);
                }else if (i == 1){
                    zhishijTable1.getColumnModel().getColumn(i).setPreferredWidth(140);
                }else {
                    zhishijTable1.getColumnModel().getColumn(i).setPreferredWidth(120);
                }

            }
        }else if (tableName.equals("chengyuanjTable1")){
            for (int i = 0; i < Constant.ChengYuan; i++) {
                chengyuanjTable1.getColumnModel().getColumn(i).setPreferredWidth(120);
            }
        }else if (tableName.equals("zuzhijigoujTable1")){
            for (int i = 0; i < Constant.ZuZhiJiGou; i++) {
                zuzhijigoujTable1.getColumnModel().getColumn(i).setPreferredWidth(132);
            }
        }else if (tableName.equals("shizhijTable1")){
            for (int i = 0; i < Constant.ShiZhiGanBu; i++) {
                if (i <= 3){
                    shizhijTable1.getColumnModel().getColumn(i).setPreferredWidth(132);
                }else {
                    shizhijTable1.getColumnModel().getColumn(i).setPreferredWidth(159);
                }
            }
        }else if (tableName.equals("zhilianjTable1")){
            for (int i = 0; i < Constant.ZhiLian; i++) {
                zhilianjTable1.getColumnModel().getColumn(i).setPreferredWidth(145);
            }
        }else {
            for (int i = 0; i < Constant.ZuZhiFaZhan; i++) {
                zuhifazhanTable1.getColumnModel().getColumn(i).setPreferredWidth(135);
            }
        }
    }


}
