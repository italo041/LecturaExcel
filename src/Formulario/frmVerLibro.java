package Formulario;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.Iterator;
import javax.swing.JOptionPane;
import javax.swing.JTable;
import javax.swing.table.DefaultTableModel;
 
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
/**
 *
 * @author USP
 */
public class frmVerLibro extends javax.swing.JDialog {

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

//        try (FileInputStream file = new FileInputStream(new File(rutaArchivo))) {
//                // leer archivo excel
//                XSSFWorkbook worbook = new XSSFWorkbook(file);
//                //obtener la hoja que se va leer
//                XSSFSheet sheet = worbook.getSheetAt(0);
//                //obtener todas las filas de la hoja excel
//                Iterator<Row> rowIterator = sheet.iterator();
//
//                Row row;
//                // se recorre cada fila hasta el final
//                while (rowIterator.hasNext()) {
//                        row = rowIterator.next();
//                        //se obtiene las celdas por fila
//                        Iterator<Cell> cellIterator = row.cellIterator();
//                        Cell cell;
//                        //se recorre cada celda
//                        while (cellIterator.hasNext()) {
//                                // se obtiene la celda en específico y se la imprime
//                                cell = cellIterator.next();
//                                System.out.print(cell.getStringCellValue()+" | ");
//                        }
//                        System.out.println();
//                }
//        } catch (Exception e) {
//                e.getMessage();
//        }
        JTable jtable= new JTable();
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
                    tableModel.addColumn("Col." + i);
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

        javax.swing.GroupLayout layout = new javax.swing.GroupLayout(getContentPane());
        getContentPane().setLayout(layout);
        layout.setHorizontalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(layout.createSequentialGroup()
                .addContainerGap()
                .addComponent(jScrollPane1, javax.swing.GroupLayout.DEFAULT_SIZE, 630, Short.MAX_VALUE)
                .addContainerGap())
        );
        layout.setVerticalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(layout.createSequentialGroup()
                .addContainerGap()
                .addComponent(jScrollPane1, javax.swing.GroupLayout.DEFAULT_SIZE, 424, Short.MAX_VALUE)
                .addContainerGap())
        );

        pack();
    }// </editor-fold>//GEN-END:initComponents

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
    private javax.swing.JScrollPane jScrollPane1;
    private javax.swing.JTable jTable1;
    // End of variables declaration//GEN-END:variables
}
