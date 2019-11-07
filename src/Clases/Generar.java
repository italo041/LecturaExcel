/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package Clases;

import java.io.File;
import java.io.IOException;
import java.util.logging.Level;
import java.util.logging.Logger;
import javafx.beans.value.WritableBooleanValue;
import jxl.Workbook;
import jxl.WorkbookSettings;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;


/**
 *
 * @author RONALDO
 */
public class Generar {
    public void generarExcel (String [][] entrada, String ruta){
    
        try {
            WorkbookSettings conf=new WorkbookSettings();
            conf.setEncoding("ISO-8859-1");        
            WritableWorkbook woorBook=  Workbook.createWorkbook(new File(ruta),conf);
            
            WritableSheet shett= woorBook.createSheet("resultado", 0);
            
            for (int i = 0; i < entrada.length; i++) {
                for (int j = 0; j < entrada[i].length; j++) {
                    try {
                        shett.addCell(new jxl.write.Label(j, i, entrada[i][j]));
                    } catch (Exception ex) {
                        Logger.getLogger(Generar.class.getName()).log(Level.SEVERE, null, ex);
                }
            }
            }
            woorBook.write();
            try {
                
                woorBook.close();
            } catch (WriteException ex) {
                Logger.getLogger(Generar.class.getName()).log(Level.SEVERE, null, ex);
            }
        } catch (IOException e) {
            Logger.getLogger(Generar.class.getName()).log(Level.SEVERE, null, e);
        }
            
      
    }
    
}
