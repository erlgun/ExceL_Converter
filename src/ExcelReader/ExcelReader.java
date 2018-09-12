/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package ExcelReader;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.poifs.filesystem.*;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import static java.lang.System.*;
import org.apache.poi.ss.usermodel.Row;
import java.io.FileOutputStream;
import org.apache.poi.ss.usermodel.*;
import java.text.NumberFormat;
import javax.imageio.ImageIO;

/**
 *
 * @author erWulf
 */
/**
 * This example shows how to use the event API for reading a file.
 */
public class ExcelReader
        {
    
    
    public File readExcel(File file) {
try {
    POIFSFileSystem fs = new POIFSFileSystem(new FileInputStream(file));
    HSSFWorkbook wb = new HSSFWorkbook(fs);
    HSSFSheet sheet = wb.getSheetAt(0);
    HSSFRow row;
    HSSFCell cell;
    HSSFCell checker;
    
    File convert;
    
    
    FileOutputStream fop = null;
    
    DataFormatter formatter = new DataFormatter();
    NumberFormat numberFormat = NumberFormat.getInstance();
    
    short x;
    x = wb.createDataFormat().getFormat("#########.##");
    CellStyle dateCellFormat = wb.createCellStyle();
    dateCellFormat.setDataFormat(x);
                        
    
    
    
    

    
    int rows; // No of rows
    rows = sheet.getPhysicalNumberOfRows();

    int cols = 0; // No of columns
    int tmp = 0;

    // This trick ensures that we get the data properly even if it doesn't start from first few rows
    for(int i = 0; i < 10 || i < rows; i++) {
        row = sheet.getRow(i);
        if(row != null) {
            tmp = sheet.getRow(i).getPhysicalNumberOfCells();
            if(tmp > cols) cols = tmp;
        }
    }
    
    try {
    
    convert = File.createTempFile("prefix-", "-suffix");
    
    fop = new FileOutputStream(convert);
      if (!convert.exists()) {
	     convert.createTempFile("prefix-", "-suffix");
             
	  }

    
        
    for(int r = 2; r < rows; r++) {
        row = sheet.getRow(r);
        if(row != null) {
            for(int c = 0; c < cols; c++) {
                checker = row.getCell((short)5, Row.RETURN_BLANK_AS_NULL);
                if (checker != null) {
                cell = row.getCell((short)c);
                if (c == 5) c = 6;
                if (c == 0) c = 4;
                
                if(cell != null) {
                    String value = null;
                    
                    if (c == 4) {

                        value = formatter.formatCellValue(cell);
                    }
                    else if (c == 6) {
                        cell.setCellStyle(dateCellFormat);
                        value = "|Z" + formatter.formatCellValue(cell);
                        value = value.replace(".","");
                    }
                    else {
                        cell.setCellStyle(dateCellFormat);
                        value = "|Z" + formatter.formatCellValue(cell) + System.getProperty("line.separator");
                        value = value.replace(".","");
                    }
                    
                    
                    byte[] contentInBytes = value.getBytes();
                    fop.write(contentInBytes);
                    fop.flush();
                    
                    
                }
                }
            }
            
        }
    }
    
    fop.close();
    out.println(convert); 
    }
    catch(Exception ioe) {
    return null;
}
    convert.deleteOnExit();
    return convert;
    
    
} catch(Exception ioe) {
    return null;
}


}
}
