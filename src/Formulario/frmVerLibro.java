package Formulario;
import java.io.BufferedWriter;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.util.Iterator;
import java.util.logging.Level;
import java.util.logging.Logger;
import javax.swing.JOptionPane;
import javax.swing.JTable;
import javax.swing.RowFilter;
import javax.swing.table.DefaultTableModel;
import javax.swing.table.TableModel;
import javax.swing.table.TableRowSorter;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
 
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
/**
 *
 * @author USP
 */
public class frmVerLibro extends javax.swing.JDialog {
    private TableRowSorter<TableModel> modeloOrdenado;
    /**
     * Creates new form frmVerLibro
     */
    public frmVerLibro(java.awt.Frame parent, boolean modal) {
        super(parent, modal);
        initComponents();
        this.setLocationRelativeTo(null);
        LeerLibro();
    }
    
    public void LeerLibro(){
        String nombreArchivo = "archivo.xlsx";
        String rutaArchivo = "libro/" + nombreArchivo;
        String [] Cabecera=new String[1];
        //************************COMBO*****************************************//
        try (FileInputStream file = new FileInputStream(new File(rutaArchivo))) {
                // leer archivo excel
                XSSFWorkbook worbook = new XSSFWorkbook(file);
                //obtener la hoja que se va leer
                XSSFSheet sheet = worbook.getSheetAt(0);
                //obtener todas las filas de la hoja excel
                Iterator<Row> rowIterator = sheet.iterator();

                Row row;
                //obtiene cantidad total de columnas con contenido
                int maxCol = 0;
                for (int a = 0; a <= sheet.getLastRowNum(); a++) {
                    if(sheet.getRow(a)!=null){
                        if (sheet.getRow(a).getLastCellNum() > maxCol) {
                            maxCol = sheet.getRow(a).getLastCellNum();
                        }    
                    }                
                }
                Cabecera= new String[maxCol];
                int col=0;
                // se recorre cada fila hasta el final
                while (rowIterator.hasNext()) {
                        row = rowIterator.next();
                        //se obtiene las celdas por fila
                        Iterator<Cell> cellIterator = row.cellIterator();
                        Cell cell;
                        //se recorre cada celda
                        while (cellIterator.hasNext()) {
                                // se obtiene la celda en específico y se la imprime
                                cell = cellIterator.next();
                                //System.out.print(cell.getStringCellValue()+" | ");
                                cboFiltrar.addItem(cell.getStringCellValue());
                                Cabecera[col]=cell.getStringCellValue();
                                col+=1;
                        }
//                        System.out.println();
                }
        } catch (Exception e) {
                e.getMessage();
        }
        //************************TABLA*****************************************//
        DefaultTableModel tableModel = new DefaultTableModel();
        try {
            XSSFWorkbook wb = new XSSFWorkbook(new FileInputStream(new File(rutaArchivo)));
            XSSFSheet sheet = wb.getSheetAt(0);//primeta hoja            
            Row row;
            Cell cell;

            //obtiene cantidad total de columnas con contenido
            int maxCol = 0;
            for (int a = 0; a <= sheet.getLastRowNum(); a++) {
                if(sheet.getRow(a)!=null){
                    if (sheet.getRow(a).getLastCellNum() > maxCol) {
                        maxCol = sheet.getRow(a).getLastCellNum();
                    }    
                }                
            }
            if (maxCol > 0) {
                //Añade encabezado a la tabla
                for (int i = 1; i <= maxCol; i++) {
                    tableModel.addColumn(Cabecera[i-1]);
                }                
                //recorre fila por fila
                Iterator<Row> rowIterator = sheet.iterator();
                while (rowIterator.hasNext()) {

                    int index = 0;
                    row = rowIterator.next();

                    Object[] obj = new Object[row.getLastCellNum()];
                    Iterator<Cell> cellIterator = row.cellIterator();

                    while (cellIterator.hasNext()) {
                        cell = cellIterator.next();
                        //contenido para celdas vacias
                        while (index < cell.getColumnIndex()) {
                            obj[index] = "";
                            index += 1;
                        }
                        
                        //extrae contenido de archivo excel
                        switch (cell.getCellType()) {
                            case BOOLEAN:
                                obj[index] = cell.getBooleanCellValue();
                                break;
                            case NUMERIC:
                                obj[index] = cell.getNumericCellValue();
                                break;
                            case STRING:
                                obj[index] = cell.getStringCellValue();
                                break;
                            case BLANK:
                                obj[index] = " ";
                                break;
                            case FORMULA:
                                obj[index] = cell.getCellFormula();
                                break;                           
                            default:
                                obj[index] = "";
                                break;
                        }                        
                        index += 1;
                    }
                    tableModel.addRow(obj);
                }
                tableModel.removeRow(0);
                jTable1.setModel(tableModel);
            }else{
                JOptionPane.showMessageDialog(null, "Nada que importar", "Error", JOptionPane.ERROR_MESSAGE);
            }
        } catch (IOException ex) {
            System.err.println("" + ex.getMessage());
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

        jScrollPane1 = new javax.swing.JScrollPane();
        jTable1 = new javax.swing.JTable();
        jLabel1 = new javax.swing.JLabel();
        cboFiltrar = new javax.swing.JComboBox<>();
        txtFiltrar = new javax.swing.JTextField();
        btnEliminar = new javax.swing.JButton();
        btnCopiar = new javax.swing.JButton();

        setDefaultCloseOperation(javax.swing.WindowConstants.DISPOSE_ON_CLOSE);

        jTable1.setModel(new javax.swing.table.DefaultTableModel(
            new Object [][] {
                {null, null, null, null},
                {null, null, null, null},
                {null, null, null, null},
                {null, null, null, null}
            },
            new String [] {
                "Title 1", "Title 2", "Title 3", "Title 4"
            }
        ));
        jTable1.setAutoResizeMode(javax.swing.JTable.AUTO_RESIZE_ALL_COLUMNS);
        jScrollPane1.setViewportView(jTable1);

        jLabel1.setText("Filtrar por: ");

        txtFiltrar.addKeyListener(new java.awt.event.KeyAdapter() {
            public void keyPressed(java.awt.event.KeyEvent evt) {
                txtFiltrarKeyPressed(evt);
            }
            public void keyReleased(java.awt.event.KeyEvent evt) {
                txtFiltrarKeyReleased(evt);
            }
        });

        btnEliminar.setText("Eliminar");
        btnEliminar.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                btnEliminarActionPerformed(evt);
            }
        });

        btnCopiar.setText("Copiar en TXT");
        btnCopiar.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                btnCopiarActionPerformed(evt);
            }
        });

        javax.swing.GroupLayout layout = new javax.swing.GroupLayout(getContentPane());
        getContentPane().setLayout(layout);
        layout.setHorizontalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(layout.createSequentialGroup()
                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addGroup(layout.createSequentialGroup()
                        .addContainerGap()
                        .addComponent(jScrollPane1, javax.swing.GroupLayout.DEFAULT_SIZE, 732, Short.MAX_VALUE))
                    .addGroup(layout.createSequentialGroup()
                        .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                            .addGroup(layout.createSequentialGroup()
                                .addGap(30, 30, 30)
                                .addComponent(jLabel1)
                                .addGap(18, 18, 18)
                                .addComponent(cboFiltrar, javax.swing.GroupLayout.PREFERRED_SIZE, 122, javax.swing.GroupLayout.PREFERRED_SIZE)
                                .addGap(18, 18, 18)
                                .addComponent(txtFiltrar, javax.swing.GroupLayout.PREFERRED_SIZE, 122, javax.swing.GroupLayout.PREFERRED_SIZE))
                            .addGroup(layout.createSequentialGroup()
                                .addGap(291, 291, 291)
                                .addComponent(btnEliminar)
                                .addGap(78, 78, 78)
                                .addComponent(btnCopiar)))
                        .addGap(0, 0, Short.MAX_VALUE)))
                .addContainerGap())
        );
        layout.setVerticalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(javax.swing.GroupLayout.Alignment.TRAILING, layout.createSequentialGroup()
                .addContainerGap(27, Short.MAX_VALUE)
                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(jLabel1)
                    .addComponent(cboFiltrar, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addComponent(txtFiltrar, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE))
                .addGap(18, 18, 18)
                .addComponent(jScrollPane1, javax.swing.GroupLayout.PREFERRED_SIZE, 403, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(btnEliminar)
                    .addComponent(btnCopiar))
                .addContainerGap())
        );

        pack();
    }// </editor-fold>//GEN-END:initComponents

    private void btnCopiarActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_btnCopiarActionPerformed
        int fila = jTable1.getSelectedRow();
        System.out.println(fila);
        try {
            String cadena = jTable1.getValueAt(fila, 6)+"|"+jTable1.getValueAt(fila, 4)+"|"+jTable1.getValueAt(fila, 5)+"|"+jTable1.getValueAt(fila, 10)+"|"+jTable1.getValueAt(fila, 9);
            System.out.println(cadena);
            String ruta = "txt/libro.txt";
            File file = new File(ruta);
            // Si el archivo no existe es creado
            if (!file.exists()) {
                file.createNewFile();
            }
            FileWriter fw = new FileWriter(file);
            BufferedWriter bw = new BufferedWriter(fw);
            bw.write(cadena);
            bw.close();
            JOptionPane.showMessageDialog(this, "Libro copiado en documento, listo para enviar");
        } catch (Exception e) {
            JOptionPane.showMessageDialog(this, "Seleccione un libro");
        }
    }//GEN-LAST:event_btnCopiarActionPerformed

    private void txtFiltrarKeyPressed(java.awt.event.KeyEvent evt) {//GEN-FIRST:event_txtFiltrarKeyPressed
        
    }//GEN-LAST:event_txtFiltrarKeyPressed

    private void txtFiltrarKeyReleased(java.awt.event.KeyEvent evt) {//GEN-FIRST:event_txtFiltrarKeyReleased
        int columna=cboFiltrar.getSelectedIndex();
        String texto=txtFiltrar.getText();
        TableModel tableModel =jTable1.getModel();
        modeloOrdenado = new TableRowSorter<TableModel>(tableModel);
        jTable1.setRowSorter(modeloOrdenado);
        modeloOrdenado.setRowFilter(RowFilter.regexFilter(texto, columna));
        jTable1.setModel(tableModel);
    }//GEN-LAST:event_txtFiltrarKeyReleased

    private void btnEliminarActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_btnEliminarActionPerformed
        String nombreArchivo = "archivo.xlsx";
        String rutaArchivo = "libro/" + nombreArchivo;
        int fila = jTable1.getSelectedRow();
        int indice =(int) Double.parseDouble(jTable1.getValueAt(fila, 0).toString());
        //JOptionPane.showMessageDialog(this, fila+" - "+indice);
        try {
            boolean rsp=deleteRow(rutaArchivo, fila+1);
            DefaultTableModel tableModel =(DefaultTableModel)jTable1.getModel();
            tableModel.removeRow(fila);
            jTable1.setModel(tableModel);
            JOptionPane.showMessageDialog(this, "Eliminado");
        } catch (IOException ex) {
            
        }
    }//GEN-LAST:event_btnEliminarActionPerformed
   
    public boolean deleteRow(String excelPath, int rowNo) throws IOException { 

        XSSFWorkbook workbook = null; 
        XSSFSheet sheet = null; 

        try { 
         FileInputStream file = new FileInputStream(new File(excelPath)); 
         workbook = new XSSFWorkbook(file); 
         sheet = workbook.getSheetAt(0); 
         if (sheet == null) { 
          return false; 
         } 
         int lastRowNum = sheet.getLastRowNum(); 
         if (rowNo >= 0 && rowNo < lastRowNum) { 
          sheet.shiftRows(rowNo + 1, lastRowNum, -1); 
         } 
         if (rowNo == lastRowNum) { 
          XSSFRow removingRow=sheet.getRow(rowNo); 
          if(removingRow != null) { 
           sheet.removeRow(removingRow); 
          } 
         } 
         file.close(); 
         FileOutputStream outFile = new FileOutputStream(new File(excelPath)); 
         workbook.write(outFile); 
         outFile.close(); 


        } catch(Exception e) { 
         throw e; 
        } finally { 
         if(workbook != null) 
          workbook.close(); 
        } 
        return false; 
    }
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
            java.util.logging.Logger.getLogger(frmVerLibro.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (InstantiationException ex) {
            java.util.logging.Logger.getLogger(frmVerLibro.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (IllegalAccessException ex) {
            java.util.logging.Logger.getLogger(frmVerLibro.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (javax.swing.UnsupportedLookAndFeelException ex) {
            java.util.logging.Logger.getLogger(frmVerLibro.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        }
        //</editor-fold>

        /* Create and display the dialog */
        java.awt.EventQueue.invokeLater(new Runnable() {
            public void run() {
                frmVerLibro dialog = new frmVerLibro(new javax.swing.JFrame(), true);
                dialog.addWindowListener(new java.awt.event.WindowAdapter() {
                    @Override
                    public void windowClosing(java.awt.event.WindowEvent e) {
                        System.exit(0);
                    }
                });
                dialog.setVisible(true);
            }
        });
    }

    // Variables declaration - do not modify//GEN-BEGIN:variables
    private javax.swing.JButton btnCopiar;
    private javax.swing.JButton btnEliminar;
    private javax.swing.JComboBox<String> cboFiltrar;
    private javax.swing.JLabel jLabel1;
    private javax.swing.JScrollPane jScrollPane1;
    private javax.swing.JTable jTable1;
    private javax.swing.JTextField txtFiltrar;
    // End of variables declaration//GEN-END:variables
}
