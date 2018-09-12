/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package excel.converter;
import gui.GuiFinal;
import javax.swing.SwingUtilities;
import javax.swing.*;

/**
 *
 * @author erWulf
 */
public class ExcelConverter {

    /**
     * @param args the command line arguments
     */
    public static void main(String[] args) {
               //Schedule a job for the event dispatch thread:
        //creating and showing this application's GUI.
        SwingUtilities.invokeLater(() -> {
            //Turn off metal's use of bold fonts
            UIManager.put("swing.boldMetal", Boolean.FALSE);
            gui.GuiFinal.createAndShowGUI();
        });
    }
    
}
