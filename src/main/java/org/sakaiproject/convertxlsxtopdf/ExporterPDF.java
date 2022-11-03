/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package org.sakaiproject.convertxlsxtopdf;

/**
 *
 * @author joseluis.caamal
 */

import java.io.FileInputStream;
import java.io.*;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.ss.usermodel.*;

import java.util.Iterator;
import com.itextpdf.text.*;
import com.itextpdf.text.pdf.*;


public class ExporterPDF {  
public static void main(String[] args) throws Exception{

        try (FileInputStream input_document = new FileInputStream(new File("C:\\reporte.xlsx"))) {
            XSSFWorkbook my_xls_workbook = new XSSFWorkbook(input_document);
            XSSFSheet my_worksheet = my_xls_workbook.getSheetAt(0);
            Iterator<Row> rowIterator = my_worksheet.iterator();
            
            Document iText_xls_2_pdf = new Document();
            PdfWriter.getInstance(iText_xls_2_pdf, new FileOutputStream("Excel2PDF_Output.pdf"));
            iText_xls_2_pdf.open();
            
            PdfPTable my_table = new PdfPTable(2);
            PdfPCell table_cell;
            while(rowIterator.hasNext()) {
                Row row = rowIterator.next();
                Iterator<Cell> cellIterator = row.cellIterator();
                while(cellIterator.hasNext()) {
                    Cell cell = cellIterator.next(); 
                    switch(cell.getCellType()) { 
                        case STRING:

                            table_cell=new PdfPCell(new Phrase(cell.getStringCellValue()));

                            my_table.addCell(table_cell);
                            break;
                    }
                }

            }
            iText_xls_2_pdf.add(my_table);
            iText_xls_2_pdf.close();
        }
            
    }
    
}

