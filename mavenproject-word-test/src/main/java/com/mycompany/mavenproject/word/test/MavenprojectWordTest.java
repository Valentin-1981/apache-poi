/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 */

package com.mycompany.mavenproject.word.test;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;
import java.util.function.Consumer;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xwpf.extractor.XWPFWordExtractor;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;

/**
 *
 * @author valentin
 */
public class MavenprojectWordTest {

    public static void main(String[] args) throws FileNotFoundException, IOException, InvalidFormatException {
        System.out.println("Hello World!");
        
        XWPFDocument document = new XWPFDocument();
        // create a new file
        FileOutputStream out = new FileOutputStream(new File("/home/valentin/Документы/document.docx"));
        // create a new paragraph paragraph
        XWPFParagraph paragraph = document.createParagraph();
        XWPFRun run = paragraph.createRun();
        run.setText("File Format Developer Guide - " +
          "Learn about computer files that you come across in " +
          "your daily work at: www.fileformat.com ");
        String text = run.getText(0);
        if(text != null && text.contains("computer")){
            text = text.replace("computer", "computer_2");
            run.setText(text, 0);
        }
//        document.write(out);
//        out.close();
        
        
//        File file;
//        file = new File("/home/valentin/Документы/document.docx");
//        FileInputStream fis = new FileInputStream(file.getAbsolutePath());
//        XWPFDocument document = new XWPFDocument(fis); // Вот и объект описанного нами класса
//        String documentLine = document.getDocument().toString(); 
//        System.out.println(documentLine);
//        document.getParagraphs();
//        XWPFParagraph lastParagraph = document.createParagraph();
//        System.out.println(lastParagraph);
//        document.createTable().createRow().createCell().addParagraph();
//        XWPFTable table = document.createTable(); //Здесь всё просто, создаем таблицу в документе и работаем с ней.
//        XWPFTableCell cell;//Добавим к таблице ряд, к ряду - ячейку, и используем её.
//        cell = table.createRow().createCell();
//        XWPFTable innerTable = new XWPFTable(cell.getCTTc().addNewTbl(), cell, 2, 2); // Воспользуемся конструктором для добавления таблицы - возьмем cell и её внутренние свойства, а так же зададим число рядов и колонок вложенной таблицы
//        cell.insertTable(cell.getTables().size(), innerTable);
//////        paragraph.createRun();
////        document = new XWPFDocument();
//        table = document.createTable(2, 2);
////        XWPFParagraph paragraph = document.createParagraph();
//        MavenprojectWordTest.fillTable(table);
//        MavenprojectWordTest.fillParagraph(paragraph);
        
        XWPFTable table = document.createTable();
        // create first row
        XWPFTableRow tableRowOne = table.getRow(0);
        tableRowOne.getCell(0).setText("Serial No");
        tableRowOne.addNewTableCell().setText("Products");
        tableRowOne.addNewTableCell().setText("Formats");
        // create second row
        XWPFTableRow tableRowTwo = table.createRow();
        tableRowTwo.getCell(0).setText("1");
        tableRowTwo.getCell(1).setText("Apache POI XWPF");
        tableRowTwo.getCell(2).setText("DOCX, HTML, FO, TXT, PDF");
        // create third row
        XWPFTableRow tableRowThree = table.createRow();
        tableRowThree.getCell(0).setText("2");
        tableRowThree.getCell(1).setText("Apache POI HWPF");
        tableRowThree.getCell(2).setText("DOC, HTML, FO, TXT");
        document.write(out);

        FileInputStream fis = new FileInputStream("/home/valentin/Документы/document.docx");
        // open file
        XWPFDocument file  = new XWPFDocument(OPCPackage.open(fis));
        // read text
        XWPFWordExtractor ext = new XWPFWordExtractor(file);
        // display text
        System.out.println(ext.getText());
        
        out.close();
        

    }
    
    @SuppressWarnings("static-access")
    public static void fillParagraph(XWPFParagraph paragraf) {
        paragraf.setIndentFromLeft(20);
        XWPFRun run = paragraf.createRun();
        run.setFontSize(12);
        run.setFontFamily("Times New Roman");
        run.setText("My text");
        run.addBreak();
        run.setText("New line");
      }
    
    public static void fillTable(XWPFTable table) {
        XWPFTableRow firstRow = table.getRows().get(0);
        XWPFTableRow secondRow = table.getRows().get(1);
        XWPFTableRow thirdRow = table.createRow();
        MavenprojectWordTest.fillRow(firstRow);
    }
    
    public static void fillRow(XWPFTableRow row) {
      List<XWPFTableCell> cellsList;
        cellsList = row.getTableCells();
      cellsList.forEach(cell -> fillParagraph(cell.addParagraph()));
   }
}


