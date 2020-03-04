/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package learnexcel;

/**
 *
 * @author Женя
 */
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.FileSystems;
import java.nio.file.Path;
import java.util.HashMap;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
public class ExcelEditor {
    HashMap<String, double[]> MyExport = new HashMap<>();
    public void export() throws FileNotFoundException, IOException {
        Path file_path = FileSystems.getDefault().getPath("Files/D32.xlsx");
        
        XSSFWorkbook MyBook = new XSSFWorkbook(new FileInputStream(file_path.toString()));
        XSSFSheet MySheet = MyBook.getSheetAt(4);
        int rowCount = MySheet.getPhysicalNumberOfRows();
        XSSFRow headers = MySheet.getRow(0);
        for (int i = 0; i < headers.getPhysicalNumberOfCells(); i++) {
            XSSFCell header = headers.getCell(i);
            String ColName = header.getStringCellValue();
            double[] values = new double[rowCount-1];
            int k=0;
            for (int j = 1; j < rowCount; j++) {
                values[k] = MySheet.getRow(j).getCell(i).getNumericCellValue();
                k++;
            }
            MyExport.put(ColName, values);
            }
        Frame.ExportLabel.setVisible(true);
    }

    void createNewBook() throws IOException {
        Workbook MyWB = new XSSFWorkbook();
        int n=MyExport.get("X").length;
        double Mx=0,My=0,Mz=0;
        double sum=0;
        for (double value : MyExport.get("X")) {
            sum+=value;
        } 
        Mx=sum/n;
        sum=0;
        for (double value : MyExport.get("Y")) {
            sum+=value;
        }
        My=sum/n;
        sum=0;
        for (double value : MyExport.get("Z")) {
            sum+=value;
        }
        Mz=sum/n;
        double covXY=0,covXZ=0,covYZ=0;
        sum=0;
        double valX=0,valY=0,valZ=0;
        for (int i=0;i<n;i++){
            valX = MyExport.get("X")[i] - Mx;
            valY = MyExport.get("Y")[i] - My;
            valZ = MyExport.get("Z")[i] - Mz;
            covXY+=valX*valY;
            covXZ+=valX*valZ;
            covYZ+=valY*valZ;
        }
        covXY/=(n-1);
        covXZ/=(n-1);
        covYZ/=(n-1);
        Sheet MyFirstSheet = MyWB.createSheet("Первый лист");
        Row MyFirstRow = MyFirstSheet.createRow(0);
        Row MySecondRow = MyFirstSheet.createRow(0);
        Row MyThirdRow = MyFirstSheet.createRow(0);
        Cell name1 = MyFirstRow.createCell(0);
        Cell value1 = MyFirstRow.createCell(1);
        Cell name2 = MySecondRow.createCell(0);
        Cell value2 = MySecondRow.createCell(1);
        Cell name3 = MyThirdRow.createCell(0);
        Cell value3 = MyThirdRow.createCell(1);
        name1.setCellValue("Ковариация XY");
        name2.setCellValue("Ковариация XZ");
        name3.setCellValue("Ковариация YZ");
        value1.setCellValue(covXY);
        value2.setCellValue(covXZ);
        value3.setCellValue(covYZ);
        Path file_path = FileSystems.getDefault().getPath("FirstTry.xlsx");
        FileOutputStream stream = new FileOutputStream(new File(file_path.toString()));
        MyWB.write(stream);
        MyWB.close();
        
    }  
    
}
