package com.doughtnerd.pod.word;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;
import java.util.regex.Pattern;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;

public final class WordOperator {
	
	//private static Pattern variableSyntax = Pattern.compile("#{\\w*}");
	private static String variableSyntax  = "#{%1$s}";
	
	public static XWPFDocument loadDocument(File inputFile) throws FileNotFoundException, IOException{
		return new XWPFDocument(new FileInputStream(inputFile));
	}
	
	public static void saveDocument(XWPFDocument doc, String outputPath) throws FileNotFoundException, IOException{
		saveDocument(doc, new File(outputPath));
	}
	
	public static void saveDocument(XWPFDocument doc, File outputFile) throws FileNotFoundException, IOException{
		doc.write(new FileOutputStream(outputFile));
	}
	
	public static void replaceInDocument(XWPFDocument doc, String lookFor, String replaceWith, boolean replaceAll){
    	for (XWPFParagraph p : doc.getParagraphs()) {
    	    List<XWPFRun> runs = p.getRuns();
    	    if (runs != null) {
    	        for (XWPFRun r : runs) {
    	        	if(searchAndReplaceInRun(r, lookFor, replaceWith) && !replaceAll){
    	        		return;
    	        	}
    	        }
    	    }
    	}
    	
    	for (XWPFTable tbl : doc.getTables()) {
    	   for (XWPFTableRow row : tbl.getRows()) {
    	      for (XWPFTableCell cell : row.getTableCells()) {
    	         for (XWPFParagraph p : cell.getParagraphs()) {
    	            for (XWPFRun r : p.getRuns()) {
    	            	if(searchAndReplaceInRun(r, lookFor, replaceWith) && !replaceAll){
    	            		return;
    	            	}
    	            }
    	         }
    	      }
    	   }
    	}
    }
    
    private static boolean searchAndReplaceInRun(XWPFRun r, String lookFor, String replace){
        String text = r.getText(0);
        lookFor = "#{"+lookFor+"}";
        if (text != null && text.contains(lookFor)) {
          text = text.replace(lookFor, replace);
          r.setText(text,0);
          return true;
        }
        return false;
    }
    
    /*
	public static Pattern getVariableSyntax() {
		return variableSyntax;
	}

	public static void setVariableSyntax(String variableSyntaxRegex) {
		WordOperator.variableSyntax = Pattern.compile(variableSyntaxRegex);
	}
	*/
}
