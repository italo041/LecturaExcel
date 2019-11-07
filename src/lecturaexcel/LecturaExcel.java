/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package lecturaexcel;

import Formulario.Principal;
import java.io.BufferedWriter;
import java.io.File;
import java.io.FileWriter;
 
/**
 *
 * @author USP
 */
public class LecturaExcel {

    /**
     * @param args the command line arguments
     */
    public static void main(String[] args) {
        Principal principal=new Principal();
        principal.setVisible(true);
         
        try {
            String ruta = "txt/libro.txt";
            String contenido = "Contenido de ejemplo";
            File file = new File(ruta);
            // Si el archivo no existe es creado
            if (!file.exists()) {
                file.createNewFile();
            }
            FileWriter fw = new FileWriter(file);
            BufferedWriter bw = new BufferedWriter(fw);
            bw.write(contenido);
            bw.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
        
        
        

        
    }
    
}
