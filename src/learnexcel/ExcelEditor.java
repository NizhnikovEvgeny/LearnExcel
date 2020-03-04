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
import org.apache.commons.math3.stat.correlation.Covariance;
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
        Path file_path = FileSystems.getDefault().getPath("Cov.xlsx");
        Covariance cov = new Covariance();
        Sheet MyFirstSheet = MyWB.createSheet("Первый лист");
        Row MyFirstRow = MyFirstSheet.createRow(0);
        Row MySecondRow = MyFirstSheet.createRow(1);
        Row MyThirdRow = MyFirstSheet.createRow(2);
        Cell name1 = MyFirstRow.createCell(0);
        Cell value1 = MyFirstRow.createCell(1);
        Cell name2 = MySecondRow.createCell(0);
        Cell value2 = MySecondRow.createCell(1);
        Cell name3 = MyThirdRow.createCell(0);
        Cell value3 = MyThirdRow.createCell(1);
        name1.setCellValue("Ковариация XY");
        name2.setCellValue("Ковариация XZ");
        name3.setCellValue("Ковариация YZ");
        value1.setCellValue(cov.covariance(MyExport.get("X"),MyExport.get("Y")));
        value2.setCellValue(cov.covariance(MyExport.get("X"),MyExport.get("Z")));
        value3.setCellValue(cov.covariance(MyExport.get("Y"),MyExport.get("Z")));
        
        FileOutputStream stream = new FileOutputStream(new File(file_path.toString()));
        MyWB.write(stream);
        MyWB.close();
        Frame.CovDone.setVisible(true);
    }  
    
}
