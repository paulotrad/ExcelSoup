package com.example;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.sl.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jsoup.Jsoup;
import org.jsoup.Connection.Method;
import org.jsoup.nodes.Document;

/**
 * @author Paul Otradovec
 * 
 * App created to fill out nomenclature for excel sheet using onlinecomponents.com,
 *
 */
public class App 
{

    static String newValue;
    static String partNumber;
    public static void main( String[] args ) throws IOException
    {
       
        FileInputStream file= new FileInputStream(new File("Q:/Projects5/Aerosonde/SUAS LAB/SPARE CONNECTORS.xlsx"));
        Workbook workbook = new XSSFWorkbook(file);
        System.out.println("Updating Column 5 with Nomenclature based of part numbers supplied using www.onlinecomponents.com");
        for(int i =0;i<7;i++){
        
        System.out.println(i+"\n\n\n\n");
            org.apache.poi.ss.usermodel.Sheet sheet = workbook.getSheetAt(i);
            
        for(Row row: sheet){
            try{
            
            partNumber=row.getCell(0).toString();
            }catch(Exception e){

            }
            //try to read cell if cell doesnt exist then create it
           
           
      
      
      
      
            if (partNumber.length()>1){
      
            
            try{
                Document doc = Jsoup.connect("https://www.onlinecomponents.com/en/keywordsearch?text="+partNumber.replace(' ','+').strip())
           
      
                .userAgent("Mozilla/5.0 (Windows; U; WindowsNT 5.1; en-US; rv1.8.1.6) Gecko/20070725 Firefox/2.0.0.6")                           
                .header ("Accept", "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8")
                .header ("accept-encoding", "gzip, deflate, sdch")
                .get();

                newValue=doc.select("p.font-size-13.d-none.d-md-block.mb-11.product-description").text();

                if(newValue.length()<1){
                    //if nt 
                    doc = Jsoup.connect("https://www.trustedparts.com/en/search/"+partNumber.replace("/","%2F").strip())
                    .userAgent("Mozilla/5.0 (Windows; U; WindowsNT 5.1; en-US; rv1.8.1.6) Gecko/20070725 Firefox/2.0.0.6")                        
                    .header ("Accept", "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8")
                    .header ("accept-encoding", "gzip, deflate, sdch")
                     .get();
                     newValue=doc.select("tbody>tr>td>span").text();
                  
                }
                
                
                System.out.println("Sheet "+i+"Row " +row.getRowNum()+"Part Number: "+partNumber + " Description: "+ newValue);
                
                 

                row.getCell(6).setCellValue(newValue);
                }catch(Exception e){
                row.createCell(6).setCellValue(newValue);
                }
           
                                       
            }

                }
                    
    
     
     
     
    }
    FileOutputStream outputStream= new FileOutputStream("Q:/Projects5/Aerosonde/SUAS LAB/SPARE CONNECTORS.xlsx");
    workbook.write(outputStream);
    workbook.close();
}
}