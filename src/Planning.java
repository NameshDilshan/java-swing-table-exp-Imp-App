import java.io.BufferedOutputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.logging.Level;
import java.util.logging.Logger;
import javax.swing.JFileChooser;
import javax.swing.JOptionPane;
import javax.swing.filechooser.FileNameExtensionFilter;
import javax.swing.table.DefaultTableModel;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Planning extends javax.swing.JFrame {

    public static final String URL = "jdbc:sqlserver://DESKTOP-T514NJC\\Batman;databaseName=Planning;portNumber=1433";
    public static final String USERNAME = "admin";
    public static final String PASSWORD = "1234";
    
    public Planning() {
        initComponents();
    }
    public Connection getConnection(){
        Connection connection = null;
        try {
            Class.forName("com.microsoft.sqlserver.jdbc.SQLServerDriver");
            connection = DriverManager.getConnection(URL, USERNAME, PASSWORD);
            System.out.println("SQL Server Connection Successfully.");
            return connection;
        }catch (ClassNotFoundException | SQLException connectionError){
            System.out.println("SQL Server Connection Problem Exception :"+ connectionError);
            return connection;
        }
    }
    
    public ArrayList<Plan> itemList(String toDate, String fromDate) throws SQLException{
        ArrayList<Plan> arrayList = new ArrayList<>();
        try{
            Connection connection = getConnection();
            String query ="SELECT " +
                    "  [FGOrder_HDR].[OrderNo], " +
                    "  [FGOrder_DTL].[ItemCode], " +
                    "  [FGOrder_DTL].[QTY], " +
                    "  [FGOrder_DTL].[QTY_Pending] AS [Completed_QTY], " +
                    "  [FGOrder_HDR].[ODate], " +
                    "  [FGOrder_HDR].[OrderedBy], " +
                    "  [FGItems].[Color], " +
                    "  [FGItems].[Weight]*[FGOrder_DTL].[QTY] AS WEIGHT, " +
                    "  [FGItems].[Weight]* [FGOrder_DTL]. [QTY_Pending] AS [Completed_WEIGHT] " +
                    " FROM " +
                    "  (( [FGOrder_DTL] " +
                    "  INNER JOIN [FGOrder_HDR] ON [FGOrder_HDR].OrderNo = [FGOrder_DTL].OrderNo) " +
                    "  INNER JOIN [FGItems] ON [FGOrder_DTL].[ItemCode]=[FGItems].[ItemCode] ) " +
                    "  WHERE [ODate] BETWEEN '"+fromDate+"' AND '"+toDate+"' ORDER BY [FGOrder_HDR].[OrderNo] ";
            Statement statement = connection.createStatement();
            ResultSet resultSet = statement.executeQuery(query);
            while(resultSet.next()){
                Plan plan = new Plan(
                        resultSet.getString("OrderNo") , 
                        resultSet.getString("ItemCode")  , 
                        resultSet.getString("QTY")  , 
                        resultSet.getString("Completed_QTY") , 
                        resultSet.getString("ODate") , 
                        resultSet.getString("OrderedBy") , 
                        resultSet.getString("Color") , 
                        resultSet.getString("WEIGHT") , 
                        resultSet.getString("Completed_WEIGHT") );
                arrayList.add(plan);
            }
        }catch(SQLException ex){
        }return arrayList;
    }
    
    public void populatePlanTable() throws SQLException{
        SimpleDateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd");
        String toDate = dateFormat.format(jDateChooserToDate.getDate());
        String fromDate = dateFormat.format(jDateChooserFromDate.getDate());
        ArrayList<Plan> plan = itemList(toDate, fromDate);
        DefaultTableModel model = (DefaultTableModel) jTable1.getModel();
        Object[] row = new Object[10];
        for(int i=0; i<plan.size(); i++ ){
            row[0] = plan.get(i).getOrderNo();
            row[1] = plan.get(i).getItemCode();
            row[2] = plan.get(i).getQTY();
            row[3] = plan.get(i).getCompleted_QTY();
            row[4] = plan.get(i).getODate();
            row[5] = plan.get(i).getOrderedBy();
            row[6] = plan.get(i).getColor();
            row[7] = plan.get(i).getWEIGHT();
            row[8] = plan.get(i).getCompleted_WEIGHT();
            model.addRow(row);
        }
    }
    
    public void exportDataToExcel() {
            FileOutputStream excelFos = null;
            XSSFWorkbook excelJTableExport = null;
            BufferedOutputStream excelBos = null;
            DefaultTableModel model = (DefaultTableModel) jTable1.getModel();
            try {
                JFileChooser excelFileChooser = new JFileChooser("C:\\Users\\Authentic\\Desktop");
                excelFileChooser.setDialogTitle("Save As ..");
                FileNameExtensionFilter fnef = new FileNameExtensionFilter("Files", "xls", "xlsx", "xlsm");
                excelFileChooser.setFileFilter(fnef);
                int chooser = excelFileChooser.showSaveDialog(null);
                if (chooser == JFileChooser.APPROVE_OPTION) {
                    excelJTableExport = new XSSFWorkbook();
                    XSSFSheet excelSheet = excelJTableExport.createSheet("Export");
                    XSSFRow row = excelSheet.createRow(0);
                    for (int j = 0; j < model.getColumnCount(); j++) {
                     XSSFCell cell = row.createCell(j);
                     cell.setCellValue(model.getColumnName(j));
                    }
                    for (int i = 0; i < model.getRowCount(); i++) {
                        XSSFRow excelRow = excelSheet.createRow(i+1);
                        for (int j = 0; j < model.getColumnCount(); j++) {
                            XSSFCell excelCell = excelRow.createCell(j);
                            excelCell.setCellValue(model.getValueAt(i, j).toString());
                        }
                    }
                    excelFos = new FileOutputStream(excelFileChooser.getSelectedFile() + ".xlsx");
                    excelBos = new BufferedOutputStream(excelFos);
                    excelJTableExport.write(excelBos);
                    JOptionPane.showMessageDialog(null, "Exported Successfully");
                }
            } catch (FileNotFoundException ex) {
                JOptionPane.showMessageDialog(null, ex);
            } catch (IOException ex) {
                JOptionPane.showMessageDialog(null, ex);
            } finally {
                try {
                    if (excelFos != null) {
                        excelFos.close();
                    }
                    if (excelBos != null) {
                        excelBos.close();
                    }
                    if (excelJTableExport != null) {
                        excelJTableExport.close();
                    }
                } catch (IOException ex) {
                    JOptionPane.showMessageDialog(null, ex);
                }
            }
        } 
    
    @SuppressWarnings("unchecked")
    // <editor-fold defaultstate="collapsed" desc="Generated Code">//GEN-BEGIN:initComponents
    private void initComponents() {

        jPanel1 = new javax.swing.JPanel();
        jLabel1 = new javax.swing.JLabel();
        jLabel2 = new javax.swing.JLabel();
        jDateChooserFromDate = new com.toedter.calendar.JDateChooser();
        jDateChooserToDate = new com.toedter.calendar.JDateChooser();
        jScrollPane1 = new javax.swing.JScrollPane();
        jTable1 = new javax.swing.JTable();
        jButtonExport = new javax.swing.JButton();
        jButtonSearchBtn = new javax.swing.JButton();

        setDefaultCloseOperation(javax.swing.WindowConstants.EXIT_ON_CLOSE);

        jPanel1.setBackground(new java.awt.Color(204, 204, 255));

        jLabel1.setFont(new java.awt.Font("Times New Roman", 0, 18)); // NOI18N
        jLabel1.setText("From");

        jLabel2.setFont(new java.awt.Font("Times New Roman", 0, 18)); // NOI18N
        jLabel2.setText("To");

        jDateChooserFromDate.setDateFormatString("yyyy-MM-dd");

        jDateChooserToDate.setDateFormatString("yyyy-MM-dd");

        jTable1.setFont(new java.awt.Font("Times New Roman", 0, 14)); // NOI18N
        jTable1.setModel(new javax.swing.table.DefaultTableModel(
            new Object [][] {

            },
            new String [] {
                "OrderNo", "ItemCode", "QTY", "Completed_QTY", "OrderDate", "OrderBy", "Color", "Weight", "Completed_Weight"
            }
        ));
        jScrollPane1.setViewportView(jTable1);

        jButtonExport.setFont(new java.awt.Font("Times New Roman", 0, 18)); // NOI18N
        jButtonExport.setText("Export");
        jButtonExport.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jButtonExportActionPerformed(evt);
            }
        });

        jButtonSearchBtn.setFont(new java.awt.Font("Times New Roman", 0, 18)); // NOI18N
        jButtonSearchBtn.setText("Search");
        jButtonSearchBtn.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jButtonSearchBtnActionPerformed(evt);
            }
        });

        javax.swing.GroupLayout jPanel1Layout = new javax.swing.GroupLayout(jPanel1);
        jPanel1.setLayout(jPanel1Layout);
        jPanel1Layout.setHorizontalGroup(
            jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(jPanel1Layout.createSequentialGroup()
                .addGap(123, 123, 123)
                .addComponent(jLabel1)
                .addGap(53, 53, 53)
                .addComponent(jDateChooserFromDate, javax.swing.GroupLayout.PREFERRED_SIZE, 163, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addGap(29, 29, 29)
                .addComponent(jLabel2)
                .addGap(31, 31, 31)
                .addComponent(jDateChooserToDate, javax.swing.GroupLayout.PREFERRED_SIZE, 162, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addGap(123, 123, 123)
                .addComponent(jButtonSearchBtn)
                .addContainerGap(javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE))
            .addGroup(javax.swing.GroupLayout.Alignment.TRAILING, jPanel1Layout.createSequentialGroup()
                .addContainerGap(javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addGroup(javax.swing.GroupLayout.Alignment.TRAILING, jPanel1Layout.createSequentialGroup()
                        .addComponent(jScrollPane1, javax.swing.GroupLayout.PREFERRED_SIZE, 1108, javax.swing.GroupLayout.PREFERRED_SIZE)
                        .addGap(23, 23, 23))
                    .addGroup(javax.swing.GroupLayout.Alignment.TRAILING, jPanel1Layout.createSequentialGroup()
                        .addComponent(jButtonExport)
                        .addGap(39, 39, 39))))
        );
        jPanel1Layout.setVerticalGroup(
            jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(jPanel1Layout.createSequentialGroup()
                .addGap(70, 70, 70)
                .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.TRAILING)
                    .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                        .addComponent(jDateChooserToDate, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                        .addComponent(jDateChooserFromDate, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                        .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                            .addComponent(jLabel1)
                            .addComponent(jLabel2)))
                    .addComponent(jButtonSearchBtn))
                .addGap(40, 40, 40)
                .addComponent(jScrollPane1, javax.swing.GroupLayout.PREFERRED_SIZE, 186, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addGap(54, 54, 54)
                .addComponent(jButtonExport)
                .addContainerGap(61, Short.MAX_VALUE))
        );

        javax.swing.GroupLayout layout = new javax.swing.GroupLayout(getContentPane());
        getContentPane().setLayout(layout);
        layout.setHorizontalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addComponent(jPanel1, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
        );
        layout.setVerticalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addComponent(jPanel1, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
        );

        pack();
    }// </editor-fold>//GEN-END:initComponents

    private void jButtonSearchBtnActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jButtonSearchBtnActionPerformed
        try {
            populatePlanTable();
        } catch (SQLException ex) {
            Logger.getLogger(Planning.class.getName()).log(Level.SEVERE, null, ex);
        }
    }//GEN-LAST:event_jButtonSearchBtnActionPerformed

    private void jButtonExportActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jButtonExportActionPerformed
          exportDataToExcel();
    }//GEN-LAST:event_jButtonExportActionPerformed

    public static void main(String args[]) {
        try {
            for (javax.swing.UIManager.LookAndFeelInfo info : javax.swing.UIManager.getInstalledLookAndFeels()) {
                if ("Nimbus".equals(info.getName())) {
                    javax.swing.UIManager.setLookAndFeel(info.getClassName());
                    break;
                }
            }
        } catch (ClassNotFoundException ex) {
            java.util.logging.Logger.getLogger(Planning.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (InstantiationException ex) {
            java.util.logging.Logger.getLogger(Planning.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (IllegalAccessException ex) {
            java.util.logging.Logger.getLogger(Planning.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (javax.swing.UnsupportedLookAndFeelException ex) {
            java.util.logging.Logger.getLogger(Planning.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        }
        
        java.awt.EventQueue.invokeLater(new Runnable() {
            public void run() {
                new Planning().setVisible(true);
            }
        });
    }

    // Variables declaration - do not modify//GEN-BEGIN:variables
    private javax.swing.JButton jButtonExport;
    private javax.swing.JButton jButtonSearchBtn;
    private com.toedter.calendar.JDateChooser jDateChooserFromDate;
    private com.toedter.calendar.JDateChooser jDateChooserToDate;
    private javax.swing.JLabel jLabel1;
    private javax.swing.JLabel jLabel2;
    private javax.swing.JPanel jPanel1;
    private javax.swing.JScrollPane jScrollPane1;
    private javax.swing.JTable jTable1;
    // End of variables declaration//GEN-END:variables
}
