package com.pranshu.exps;

import java.awt.Desktop;
import java.io.File;
import java.io.IOException;
import java.util.List;

import com.pranshu.exps.util.ExcelCreator;
import com.pranshu.exps.util.InputFileReader;

/**
 * Hello world!
 *
 */
public class App 
{
    public static void main( String[] args )
    {
    	try {
	    	String fileName = "D:/Expense/expense-input.txt";
	        String outputFileName = "D:/Expense/output-excel.xlsx";
	        List<Object> records = InputFileReader.readFile(fileName);
	        new ExcelCreator().convertRecords(records, outputFileName);
        	Desktop.getDesktop().open(new File(outputFileName));
        	System.out.println("Bhiyo done!");
		} catch (IOException e) {
			e.printStackTrace();
		}
    }
}
