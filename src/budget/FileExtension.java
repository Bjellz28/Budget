package budget;

/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */
import java.io.File;
import javax.swing.filechooser.*;

/**
 *
 * @author Bill
 */
public class FileExtension extends FileFilter {
    private String fileFormat = "xls";
    public boolean accept(File f){
        if(f.isDirectory()){
            return true;
        }
        if(extension(f).equalsIgnoreCase(fileFormat)) {
            return true;
        }else{
            return false;
        }
    }
    public String getDescription(){
        return "xls only";
    }

    public String extension(File f){
        String fileName = f.getName();
        int indexFile = fileName.lastIndexOf(".");
        if(indexFile > 0 && indexFile < fileName.length() - 1){
            return fileName.substring(indexFile+1);
            
        }else{
            return "";
        }
    }
    
}
