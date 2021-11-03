/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package quanli;

import KetNoiSQL.KetNoi;
import java.io.File;
import java.io.FileOutputStream;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.util.Vector;
import javax.swing.table.DefaultTableModel;
import javax.naming.spi.DirStateFactory;
import javax.swing.JFileChooser;
import javax.swing.JOptionPane;
import javax.swing.JTable;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
/**
 *
 * @author Minh Huan
 */
public class ThongKe extends javax.swing.JFrame {

    /**
     * Creates new form ThongKe
     */
    java.sql.Connection ketNoi = (java.sql.Connection) KetNoi.ConnectSQL();
    String manv;
    public ThongKe(String maNV) {
        initComponents();
        this.setLocationRelativeTo(null);
        layDuLieu("2021-01-01", "2021-12-31");
        jComboBox_Nam_TK.setSelectedItem("2021");
        manv = maNV;
        docLoaiTB();
    }
    private String phanquyen (String maNV){
        String sql = "select PHANQUYEN from TKNHANVIEN where MANV = ?";
        
        try {
            PreparedStatement ps = ketNoi.prepareStatement(sql);
            ps.setString(1, maNV);
            ResultSet rs = ps.executeQuery();
            if(rs.next()){
                return rs.getString("PHANQUYEN");
            }
            else {
                return "NHANVIEN";
            }
        } catch (SQLException e) {
        }
            
        return "NHANVIEN";
    }
    
    
    private void docLoaiTB(){
        String sql = "select * from LOAITB";
        jComboBox_LoaiTB_TK.removeAllItems();
        jComboBox_LoaiTB_TK.addItem("--Tất cả--");
        try {
            PreparedStatement ps = ketNoi.prepareStatement(sql);
            ResultSet rs = ps.executeQuery();
            while(rs.next()){
                String tenLoai = rs.getString("TENLOAI");
                jComboBox_LoaiTB_TK.addItem(tenLoai);
            }
        } catch (Exception e) {
            
        }
            
        
    }
    public void exportExcel(JTable table){
        
        try {
          
            JFileChooser jFile = new JFileChooser();
            jFile.showSaveDialog(this);
            File saveFile = jFile.getSelectedFile();
            if(saveFile != null) {
                saveFile = new File(saveFile.toString()+"_ThietBi.xlsx"); 
                Workbook wb = new XSSFWorkbook();
                org.apache.poi.ss.usermodel.Sheet sheetTB = wb.createSheet("THONG KE");
                Row rowCol = sheetTB.createRow(0);
                for(int i = 0; i<table.getColumnCount(); i++){
                    Cell cell = rowCol.createCell(i);
                    cell.setCellValue(table.getColumnName(i));
                }           
                for(int j=0;j < table.getRowCount(); j++){
                    Row row = sheetTB.createRow(j+1);
                    for(int k = 0; k< table.getColumnCount(); k++){
                        Cell cell = row.createCell(k);
                        if(table.getValueAt(j, k) != null){
                            cell.setCellValue(table.getValueAt(j, k).toString());
                        }
                    }
                }
                
                try (FileOutputStream f =  new FileOutputStream(saveFile)) {
                    wb.write(f);
                }
            }
            
            JOptionPane.showMessageDialog(rootPane, "Xuất file thành công");       
        } catch (Exception e) {
            e.printStackTrace();
        }
  



   }
    private  void locDuLieu(){
        String nam = (String) jComboBox_Nam_TK.getSelectedItem();
        int thang = jComboBox_Thang_TK.getSelectedIndex();
        String ngayBD = null;
        String ngayKT = null;
        if(thang == 0) {
            ngayBD = nam + "-01-01";
            ngayKT = nam + "-12-31";
        }
        else {
            int tuan = jComboBox_Tuan_TK.getSelectedIndex();
            String t = jComboBox_Thang_TK.getItemAt(thang);
            String n = traSoNgayCuaThang(t, nam);
            if(tuan == 0){
                ngayBD = nam + "-" + t + "-01";
                ngayKT = nam + "-" + t + "-"+n;
            }
            else{
                String tuanString = jComboBox_Tuan_TK.getItemAt(tuan);
                switch(tuan){
                    case 1:
                        ngayBD = nam + "-" + t + "-01";
                        ngayKT = nam + "-" + t + "-07";
                        break;
                    case 2:
                        ngayBD = nam + "-" + t + "-08";
                        ngayKT = nam + "-" + t + "-14";
                        break;
                    case 3:
                        ngayBD = nam + "-" + t + "-09";
                        ngayKT = nam + "-" + t + "-22";
                        break;
                    case 4: 
                        
                        ngayBD = nam + "-" + t + "-23";
                        ngayKT = nam + "-" + t + "-"+n;
                        break;
                }
            }
        }
        layDuLieu(ngayBD, ngayKT);
    }
    private String traSoNgayCuaThang(String thang, String nam){
        int t = Integer.parseInt(thang);
        int n = Integer.parseInt(nam);
        String ngay;
        if(t == 1 || t == 3 || t == 5 || t == 7 || t == 8 || t == 10 || t == 12){
            ngay = "31";
        }
        if(t == 4 || t == 6 || t == 9 || t == 11){
            ngay = "30";
        }
        else {
            if(n%4 == 0){
                ngay = "29";
            }
            else {
                ngay = "28";
            }
        }
        return ngay;
    }
    
    
    private void layDuLieu(String ngayBatDau, String ngayKetThuc){
        DefaultTableModel dtm = (DefaultTableModel) jTable_TK.getModel();
        dtm.setNumRows(0);
        String sql = "select CTM.MATB, COUNT(CTM.MATB) SOLAN, TB.TENTB from CHITIETMUON CTM, THIETBI TB  "
                + "where MAMUON in ( "
                + "	select MAMUON "
                + "	from MUONTB "
                + "	where NGAYGIOMUON > ? and NGAYGIODUDINHTRA < ? "
                + " ) and CTM.MATB = TB.MATB "
                + "group by "
                + "CTM.MATB, TENTB "
                + "order by SOLAN DESC";
        try {
            PreparedStatement ps = ketNoi.prepareStatement(sql);

            ps.setString(1, ngayBatDau);
            ps.setString(2, ngayKetThuc);
            ResultSet rs = ps.executeQuery();
            Vector vt;
            int i = 0;
            while (rs.next()) {
                i++;

                vt = new Vector();
                vt.add(i);
                String maTB = rs.getString("MATB");
                vt.add(maTB);
                String tenTB = rs.getString("TENTB");
                vt.add(tenTB);
                String solan = rs.getString("SOLAN");
                vt.add(solan);

                dtm.addRow(vt);
            }
            jTable_TK.setModel(dtm);
        } catch (SQLException e) {
            e.printStackTrace();
        }
    }
    

    /**
     * This method is called from within the constructor to initialize the form.
     * WARNING: Do NOT modify this code. The content of this method is always
     * regenerated by the Form Editor.
     */
    @SuppressWarnings("unchecked")
    // <editor-fold defaultstate="collapsed" desc="Generated Code">//GEN-BEGIN:initComponents
    private void initComponents() {

        jPanel1 = new javax.swing.JPanel();
        jLabel1 = new javax.swing.JLabel();
        jPanel2 = new javax.swing.JPanel();
        jScrollPane1 = new javax.swing.JScrollPane();
        jTable_TK = new javax.swing.JTable();
        jComboBox_Nam_TK = new javax.swing.JComboBox<>();
        jComboBox_Thang_TK = new javax.swing.JComboBox<>();
        jComboBox_Tuan_TK = new javax.swing.JComboBox<>();
        jLabel2 = new javax.swing.JLabel();
        jLabel3 = new javax.swing.JLabel();
        jLabel4 = new javax.swing.JLabel();
        jComboBox_LoaiTB_TK = new javax.swing.JComboBox<>();
        jLabel5 = new javax.swing.JLabel();
        jSeparator1 = new javax.swing.JSeparator();
        jLabel6 = new javax.swing.JLabel();
        jComboBox_SoTBHien_TK = new javax.swing.JComboBox<>();
        jLabel7 = new javax.swing.JLabel();
        jSeparator2 = new javax.swing.JSeparator();
        jLabel8 = new javax.swing.JLabel();
        jButton_QuayLai_TK = new javax.swing.JButton();
        jLabel9 = new javax.swing.JLabel();
        jButton_XuatFileThongKe = new javax.swing.JButton();

        setDefaultCloseOperation(javax.swing.WindowConstants.EXIT_ON_CLOSE);
        setUndecorated(true);

        jPanel1.setBackground(new java.awt.Color(255, 255, 255));

        jLabel1.setFont(new java.awt.Font("sansserif", 1, 24)); // NOI18N
        jLabel1.setForeground(new java.awt.Color(0, 0, 102));
        jLabel1.setText("THỐNG KÊ");

        jPanel2.setBackground(new java.awt.Color(0, 0, 102));

        javax.swing.GroupLayout jPanel2Layout = new javax.swing.GroupLayout(jPanel2);
        jPanel2.setLayout(jPanel2Layout);
        jPanel2Layout.setHorizontalGroup(
            jPanel2Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGap(0, 361, Short.MAX_VALUE)
        );
        jPanel2Layout.setVerticalGroup(
            jPanel2Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGap(0, 0, Short.MAX_VALUE)
        );

        jTable_TK.setModel(new javax.swing.table.DefaultTableModel(
            new Object [][] {

            },
            new String [] {
                "STT", "MATB", "TÊN TB", "SỐ LẦN MƯỢN"
            }
        ) {
            boolean[] canEdit = new boolean [] {
                false, false, false, false
            };

            public boolean isCellEditable(int rowIndex, int columnIndex) {
                return canEdit [columnIndex];
            }
        });
        jScrollPane1.setViewportView(jTable_TK);
        if (jTable_TK.getColumnModel().getColumnCount() > 0) {
            jTable_TK.getColumnModel().getColumn(0).setMinWidth(100);
            jTable_TK.getColumnModel().getColumn(0).setMaxWidth(100);
            jTable_TK.getColumnModel().getColumn(1).setResizable(false);
            jTable_TK.getColumnModel().getColumn(2).setResizable(false);
            jTable_TK.getColumnModel().getColumn(3).setMinWidth(100);
            jTable_TK.getColumnModel().getColumn(3).setMaxWidth(100);
        }

        jComboBox_Nam_TK.setModel(new javax.swing.DefaultComboBoxModel<>(new String[] { "2020", "2021", "2022", "2023" }));
        jComboBox_Nam_TK.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jComboBox_Nam_TKActionPerformed(evt);
            }
        });

        jComboBox_Thang_TK.setModel(new javax.swing.DefaultComboBoxModel<>(new String[] { "--Tất cả--", "1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12" }));
        jComboBox_Thang_TK.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jComboBox_Thang_TKActionPerformed(evt);
            }
        });

        jComboBox_Tuan_TK.setModel(new javax.swing.DefaultComboBoxModel<>(new String[] { "--Tất cả--", "1", "2", "3", "4" }));
        jComboBox_Tuan_TK.setEnabled(false);
        jComboBox_Tuan_TK.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jComboBox_Tuan_TKActionPerformed(evt);
            }
        });

        jLabel2.setText("Năm:");

        jLabel3.setText("Tháng:");

        jLabel4.setText("Tuần:");

        jLabel5.setText("Loại thiết bị");

        jLabel6.setFont(new java.awt.Font("sansserif", 1, 12)); // NOI18N
        jLabel6.setText("Thời gian");

        jComboBox_SoTBHien_TK.setModel(new javax.swing.DefaultComboBoxModel<>(new String[] { "--Tất cả--", "5", "10", "15", "20" }));

        jLabel7.setFont(new java.awt.Font("sansserif", 1, 12)); // NOI18N
        jLabel7.setText("Thuộc tính");

        jLabel8.setText("Số thiết bị hiển thị");

        jButton_QuayLai_TK.setText("Quay lại");
        jButton_QuayLai_TK.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jButton_QuayLai_TKActionPerformed(evt);
            }
        });

        jLabel9.setFont(new java.awt.Font("sansserif", 1, 18)); // NOI18N
        jLabel9.setForeground(new java.awt.Color(0, 0, 102));
        jLabel9.setText("SỐ LẦN THIẾT BỊ ĐƯỢC MƯỢN");

        jButton_XuatFileThongKe.setText("Xuất file");
        jButton_XuatFileThongKe.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jButton_XuatFileThongKeActionPerformed(evt);
            }
        });

        javax.swing.GroupLayout jPanel1Layout = new javax.swing.GroupLayout(jPanel1);
        jPanel1.setLayout(jPanel1Layout);
        jPanel1Layout.setHorizontalGroup(
            jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(javax.swing.GroupLayout.Alignment.TRAILING, jPanel1Layout.createSequentialGroup()
                .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.TRAILING)
                    .addGroup(jPanel1Layout.createSequentialGroup()
                        .addGap(58, 58, 58)
                        .addComponent(jLabel2, javax.swing.GroupLayout.PREFERRED_SIZE, 41, javax.swing.GroupLayout.PREFERRED_SIZE)
                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                        .addComponent(jComboBox_Nam_TK, javax.swing.GroupLayout.PREFERRED_SIZE, 86, javax.swing.GroupLayout.PREFERRED_SIZE)
                        .addGap(41, 41, 41)
                        .addComponent(jLabel3, javax.swing.GroupLayout.PREFERRED_SIZE, 41, javax.swing.GroupLayout.PREFERRED_SIZE)
                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                        .addComponent(jComboBox_Thang_TK, javax.swing.GroupLayout.PREFERRED_SIZE, 83, javax.swing.GroupLayout.PREFERRED_SIZE)
                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                        .addComponent(jLabel4, javax.swing.GroupLayout.PREFERRED_SIZE, 35, javax.swing.GroupLayout.PREFERRED_SIZE)
                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                        .addComponent(jComboBox_Tuan_TK, javax.swing.GroupLayout.PREFERRED_SIZE, 90, javax.swing.GroupLayout.PREFERRED_SIZE))
                    .addGroup(javax.swing.GroupLayout.Alignment.LEADING, jPanel1Layout.createSequentialGroup()
                        .addGap(42, 42, 42)
                        .addComponent(jLabel1)
                        .addGap(18, 18, 18)
                        .addComponent(jPanel2, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                        .addGap(0, 0, Short.MAX_VALUE))
                    .addGroup(jPanel1Layout.createSequentialGroup()
                        .addGap(39, 39, 39)
                        .addComponent(jLabel6, javax.swing.GroupLayout.PREFERRED_SIZE, 76, javax.swing.GroupLayout.PREFERRED_SIZE)
                        .addGap(15, 15, 15)
                        .addComponent(jSeparator1))
                    .addGroup(jPanel1Layout.createSequentialGroup()
                        .addContainerGap(javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                        .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.TRAILING)
                            .addGroup(jPanel1Layout.createSequentialGroup()
                                .addComponent(jLabel7, javax.swing.GroupLayout.PREFERRED_SIZE, 76, javax.swing.GroupLayout.PREFERRED_SIZE)
                                .addGap(18, 18, 18)
                                .addComponent(jSeparator2, javax.swing.GroupLayout.PREFERRED_SIZE, 417, javax.swing.GroupLayout.PREFERRED_SIZE))
                            .addGroup(jPanel1Layout.createSequentialGroup()
                                .addComponent(jLabel5, javax.swing.GroupLayout.PREFERRED_SIZE, 70, javax.swing.GroupLayout.PREFERRED_SIZE)
                                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                                .addComponent(jComboBox_LoaiTB_TK, javax.swing.GroupLayout.PREFERRED_SIZE, 157, javax.swing.GroupLayout.PREFERRED_SIZE)
                                .addGap(29, 29, 29)
                                .addComponent(jLabel8)
                                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                                .addComponent(jComboBox_SoTBHien_TK, javax.swing.GroupLayout.PREFERRED_SIZE, 131, javax.swing.GroupLayout.PREFERRED_SIZE))
                            .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING, false)
                                .addGroup(javax.swing.GroupLayout.Alignment.TRAILING, jPanel1Layout.createSequentialGroup()
                                    .addComponent(jButton_XuatFileThongKe, javax.swing.GroupLayout.PREFERRED_SIZE, 73, javax.swing.GroupLayout.PREFERRED_SIZE)
                                    .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                                    .addComponent(jButton_QuayLai_TK, javax.swing.GroupLayout.PREFERRED_SIZE, 100, javax.swing.GroupLayout.PREFERRED_SIZE))
                                .addComponent(jScrollPane1, javax.swing.GroupLayout.Alignment.TRAILING, javax.swing.GroupLayout.PREFERRED_SIZE, 509, javax.swing.GroupLayout.PREFERRED_SIZE)))))
                .addGap(30, 30, 30))
            .addGroup(javax.swing.GroupLayout.Alignment.TRAILING, jPanel1Layout.createSequentialGroup()
                .addGap(0, 0, Short.MAX_VALUE)
                .addComponent(jLabel9)
                .addGap(149, 149, 149))
        );
        jPanel1Layout.setVerticalGroup(
            jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(jPanel1Layout.createSequentialGroup()
                .addGap(30, 30, 30)
                .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING, false)
                    .addComponent(jLabel1, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                    .addComponent(jPanel2, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE))
                .addGap(25, 25, 25)
                .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.TRAILING)
                    .addComponent(jSeparator1, javax.swing.GroupLayout.PREFERRED_SIZE, 10, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addComponent(jLabel6))
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(jLabel2)
                    .addComponent(jLabel3)
                    .addComponent(jComboBox_Nam_TK, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addComponent(jComboBox_Thang_TK, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addComponent(jLabel4)
                    .addComponent(jComboBox_Tuan_TK, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE))
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.UNRELATED)
                .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addComponent(jLabel7, javax.swing.GroupLayout.Alignment.TRAILING)
                    .addComponent(jSeparator2, javax.swing.GroupLayout.Alignment.TRAILING, javax.swing.GroupLayout.PREFERRED_SIZE, 10, javax.swing.GroupLayout.PREFERRED_SIZE))
                .addGap(7, 7, 7)
                .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(jLabel5)
                    .addComponent(jLabel8, javax.swing.GroupLayout.PREFERRED_SIZE, 16, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addComponent(jComboBox_LoaiTB_TK, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addComponent(jComboBox_SoTBHien_TK, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE))
                .addGap(24, 24, 24)
                .addComponent(jLabel9, javax.swing.GroupLayout.PREFERRED_SIZE, 28, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addComponent(jScrollPane1, javax.swing.GroupLayout.PREFERRED_SIZE, 217, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.UNRELATED)
                .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(jButton_QuayLai_TK)
                    .addComponent(jButton_XuatFileThongKe))
                .addGap(15, 15, 15))
        );

        javax.swing.GroupLayout layout = new javax.swing.GroupLayout(getContentPane());
        getContentPane().setLayout(layout);
        layout.setHorizontalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addComponent(jPanel1, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
        );
        layout.setVerticalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addComponent(jPanel1, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
        );

        pack();
    }// </editor-fold>//GEN-END:initComponents

    private void jButton_QuayLai_TKActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jButton_QuayLai_TKActionPerformed
        // TODO add your handling code here:
        this.setVisible(false);
        new MainFrame(manv, phanquyen(manv)).setVisible(true);
    }//GEN-LAST:event_jButton_QuayLai_TKActionPerformed

    private void jComboBox_Thang_TKActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jComboBox_Thang_TKActionPerformed
        // CHon thang:
        String thang = (String) jComboBox_Thang_TK.getSelectedItem();
        if ("--Tất cả--".equals(thang)) {
            jComboBox_Tuan_TK.setEnabled(false);
        } else {
            jComboBox_Tuan_TK.setEnabled(true);
        }
        locDuLieu();
        
    }//GEN-LAST:event_jComboBox_Thang_TKActionPerformed

    private void jComboBox_Nam_TKActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jComboBox_Nam_TKActionPerformed
        // Chon nam:
        locDuLieu();
    }//GEN-LAST:event_jComboBox_Nam_TKActionPerformed

    private void jComboBox_Tuan_TKActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jComboBox_Tuan_TKActionPerformed
        // TODO add your handling code here:
        locDuLieu();
    }//GEN-LAST:event_jComboBox_Tuan_TKActionPerformed

    private void jButton_XuatFileThongKeActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jButton_XuatFileThongKeActionPerformed
        // TODO add your handling code here:
        exportExcel(jTable_TK);
    }//GEN-LAST:event_jButton_XuatFileThongKeActionPerformed

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
            java.util.logging.Logger.getLogger(ThongKe.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (InstantiationException ex) {
            java.util.logging.Logger.getLogger(ThongKe.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (IllegalAccessException ex) {
            java.util.logging.Logger.getLogger(ThongKe.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (javax.swing.UnsupportedLookAndFeelException ex) {
            java.util.logging.Logger.getLogger(ThongKe.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        }
        //</editor-fold>

        /* Create and display the form */
        java.awt.EventQueue.invokeLater(new Runnable() {
            public void run() {
                new ThongKe("QL01").setVisible(true);
            }
        });
    }

    // Variables declaration - do not modify//GEN-BEGIN:variables
    private javax.swing.JButton jButton_QuayLai_TK;
    private javax.swing.JButton jButton_XuatFileThongKe;
    private javax.swing.JComboBox<String> jComboBox_LoaiTB_TK;
    private javax.swing.JComboBox<String> jComboBox_Nam_TK;
    private javax.swing.JComboBox<String> jComboBox_SoTBHien_TK;
    private javax.swing.JComboBox<String> jComboBox_Thang_TK;
    private javax.swing.JComboBox<String> jComboBox_Tuan_TK;
    private javax.swing.JLabel jLabel1;
    private javax.swing.JLabel jLabel2;
    private javax.swing.JLabel jLabel3;
    private javax.swing.JLabel jLabel4;
    private javax.swing.JLabel jLabel5;
    private javax.swing.JLabel jLabel6;
    private javax.swing.JLabel jLabel7;
    private javax.swing.JLabel jLabel8;
    private javax.swing.JLabel jLabel9;
    private javax.swing.JPanel jPanel1;
    private javax.swing.JPanel jPanel2;
    private javax.swing.JScrollPane jScrollPane1;
    private javax.swing.JSeparator jSeparator1;
    private javax.swing.JSeparator jSeparator2;
    private javax.swing.JTable jTable_TK;
    // End of variables declaration//GEN-END:variables
}
