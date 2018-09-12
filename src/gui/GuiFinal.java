/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package gui;

/**
 *
 * @author erWulf
 */
import javax.swing.*;
import java.awt.*;
import java.awt.event.ActionEvent;
import java.awt.event.*;
import java.awt.event.ActionListener;
import javax.swing.SwingUtilities;
import javax.swing.filechooser.*;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import ExcelReader.ExcelReader;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.OutputStream;

/**
 *
 * @author erWulf
 */
public class GuiFinal extends JPanel
                             implements ActionListener {
    static private final String newline = "\n";
    JButton openButton, conButton, saveButton; 
    JTextArea log;
    JFileChooser fc;
    ExcelReader er = new ExcelReader();
    File file = null;
    File convert = null;
    File store = null;      
    
    
    
    public GuiFinal() {
        super(new BorderLayout());

        //Create the log first, because the action listeners
        //need to refer to it.
        log = new JTextArea(5,20);
        log.setMargin(new Insets(5,5,5,5));
        log.setEditable(false);
        JScrollPane logScrollPane = new JScrollPane(log);

        //Create a file chooser
        fc = new JFileChooser();
        FileNameExtensionFilter filter = new FileNameExtensionFilter("Excel file", "xls", "xlsx");
        FileNameExtensionFilter filter2 = new FileNameExtensionFilter("in file" ,"in");
        fc.addChoosableFileFilter(filter);
        fc.addChoosableFileFilter(filter2);
        fc.setAcceptAllFileFilterUsed(false);
        

        //Uncomment one of the following lines to try a different
        //file selection mode.  The first allows just directories
        //to be selected (and, at least in the Java look and feel,
        //shown).  The second allows both files and directories
        //to be selected.  If you leave these lines commented out,
        //then the default mode (FILES_ONLY) will be used.
        //
        //fc.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
        //fc.setFileSelectionMode(JFileChooser.FILES_AND_DIRECTORIES);

        //Create the open button.  We use the image from the JLF
        //Graphics Repository (but we extracted it from the jar).
        openButton = new JButton("Open a File...",
                                 createImageIcon("images/Open16.gif"));
        openButton.addActionListener(this);

        //Create the save button.  We use the image from the JLF
        //Graphics Repository (but we extracted it from the jar).
        saveButton = new JButton("Save a File...",
                                 createImageIcon("images/Save16.gif"));
        saveButton.addActionListener(this);

        //Add Convert Button
        conButton = new JButton( "Convert");
        conButton.addActionListener(this);
        
        //For layout purposes, put the buttons in a separate panel
        JPanel buttonPanel = new JPanel(); //use FlowLayout
        JPanel bottombuttonPanel = new JPanel();
        buttonPanel.add(openButton);
        buttonPanel.add(saveButton);
        bottombuttonPanel.add(conButton);
        
        

        //Add the buttons and the log to this panel.
        add(buttonPanel, BorderLayout.PAGE_START);
        add(bottombuttonPanel, BorderLayout.SOUTH);
        add(logScrollPane, BorderLayout.CENTER);
        

    }

    @Override
    public void actionPerformed(ActionEvent e) {

        //Handle open button action.
        if (e.getSource() == openButton) {
            
            int returnVal = fc.showOpenDialog(GuiFinal.this);

            
            if (returnVal == JFileChooser.APPROVE_OPTION) {
                file = fc.getSelectedFile();
                //This is where a real application would open the file.
                log.append("File Loaded." + newline);
            } else {
                log.append("Open command cancelled by user." + newline);
            }
            log.setCaretPosition(log.getDocument().getLength());

        //Handle save button action.
        } else if (e.getSource() == saveButton) {
            if (convert != null) {
            fc.setSelectedFile(new File(""));
            fc.setSelectedFiles(new File[]{new File("")});       
            int returnVal = fc.showSaveDialog(GuiFinal.this);
            if (returnVal == JFileChooser.APPROVE_OPTION) {
                
                try {
                    InputStream is = new FileInputStream(convert);
                    OutputStream os = new FileOutputStream(fc.getSelectedFile() + ".in");
                    int c;
                    while ((c = is.read()) != -1) {
                    
                        os.write(c);
                            }
                    is.close();
                    os.close();        

                //This is where a real application would save the file.
                log.append("Coverted File Saved." + newline);
                } catch (IOException ex) {
            log.append("Error saving");
        }
                
            } else {
                log.append("Save command cancelled by user." + newline);
            }
            log.setCaretPosition(log.getDocument().getLength());
            }
            else {
                log.append("Please convert a file." + newline);
            }
        } else if (e.getSource() == conButton) {
            
            if (file != null) {
            convert = er.readExcel(file);
            log.append("File Converted" + newline);
            log.setCaretPosition(log.getDocument().getLength());
            }
            else {
                log.append("Select a file to convert." + newline);
            }
        }
        
    }

    /** Returns an ImageIcon, or null if the path was invalid.
     * @param path
     * @return  */
    protected static ImageIcon createImageIcon(String path) {
        java.net.URL imgURL = GuiFinal.class.getResource(path);
        if (imgURL != null) {
            return new ImageIcon(imgURL);
        } else {
            System.err.println("Couldn't find file: " + path);
            return null;
        }
    }
    
    

    
    

    /**
     * Create the GUI and show it.  For thread safety,
     * this method should be invoked from the
     * event dispatch thread.
     */
    public static void createAndShowGUI() {
        //Create and set up the window.
        JFrame frame = new JFrame("FileChooserDemo");
        frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);

        //Add content to the window.
        frame.add(new GuiFinal());

        //Display the window.
        frame.pack();
        frame.setVisible(true);
         //make sure the program exits when the frame closes
        frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        frame.setTitle("Excel Converter");
        frame.setSize(300,250);
      
        //This will center the JFrame in the middle of the screen
        frame.setLocationRelativeTo(null);
        
        
        //The ActionListener class is used to handle the
        //event that happens when the user clicks the button.
        //As there is not a lot that needs to happen we can 
        //define an anonymous inner class to make the code simpler.
       
        //make sure the JFrame is visible
        frame.setVisible(true);
        
        
        
    }
    
    

}
