package com.pranshu.exps.util;

import java.io.FileOutputStream;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelCreator {
	
	Workbook workbook = new XSSFWorkbook();
	Sheet sheet = workbook.createSheet();
	CellStyle cellStyle = workbook.createCellStyle();
	
	public void convertRecords(List<Object> records, String outputFileName) {
		Iterator<Object> itr = records.iterator();
		int rowNum = 0, colNum = 0;
		Row dateRow = sheet.createRow(rowNum++);
		cellStyle.setAlignment(HorizontalAlignment.CENTER);
		
        while(itr.hasNext()) {
        	String line = itr.next().toString();
        	String[] lineTokens = line.split("-");
        	
        	
        	if(lineTokens.length == 1 && !line.isEmpty()) {
        		Cell cell = dateRow.createCell(colNum);
        		cell.setCellValue(line.trim());
        		cell.setCellStyle(cellStyle);
        		colNum = colNum + 2;
        		sheet.addMergedRegion(new CellRangeAddress(dateRow.getRowNum(), dateRow.getRowNum(), colNum - 2, colNum - 1));
        		//System.out.print(line.trim() + " ==> ");
        	} else if(lineTokens.length == 2) {
        		Row dataRow = sheet.getRow(rowNum++);
        		if(dataRow == null)
        			dataRow = sheet.createRow(rowNum - 1);
        		Cell item = dataRow.createCell(colNum - 2);
        		Cell price = dataRow.createCell(colNum - 1);
        		item.setCellValue(lineTokens[0].trim());
        		item.setCellStyle(cellStyle);
        		if(lineTokens[1].trim().split("-").length > 1 || lineTokens[1].trim().split("\\+").length > 1)
        			price.setCellFormula(lineTokens[1].trim());
        		else
        			price.setCellValue(Integer.valueOf(lineTokens[1].trim()));
        		price.setCellStyle(cellStyle);
        		//System.out.print(lineTokens[0].trim() + " >>> " + lineTokens[1].trim() + ", ");
        	} else if(line.isEmpty()) {
        		rowNum = 1;
        		//System.out.println();
        	} else if(line.indexOf("-") != line.lastIndexOf("-")) {
        		String[] newLine = handleNegativeBalance(line).split("`");
        		Row dataRow = sheet.getRow(rowNum++);
        		if(dataRow == null)
        			dataRow = sheet.createRow(rowNum - 1);
        		Cell item = dataRow.createCell(colNum - 2);
        		Cell price = dataRow.createCell(colNum - 1);
        		item.setCellValue(newLine[0].trim());
        		item.setCellStyle(cellStyle);
        		if(newLine[1].trim().split("-").length > 1 || newLine[1].trim().split("\\+").length > 1)
        			price.setCellFormula(newLine[1].trim());
        		else
        			price.setCellValue(Integer.valueOf(newLine[1].trim()));
        		price.setCellStyle(cellStyle);
        		//System.out.print(newLine[0].trim() + " >>> " + newLine[1].trim() + ", ");
        	}
        }
        
        autoSizeColumns();
        
        try {
			FileOutputStream fileOut = new FileOutputStream(outputFileName);
	        workbook.write(fileOut);
	        fileOut.close();
	
	        // Closing the workbook
	        workbook.close();
		} catch(Exception ex) {}
	}
	
	public void autoSizeColumns() {
	    int numberOfSheets = workbook.getNumberOfSheets();
	    for (int i = 0; i < numberOfSheets; i++) {
	        Sheet sheet = workbook.getSheetAt(i);
	        if (sheet.getPhysicalNumberOfRows() > 0) {
	            Row row = sheet.getRow(sheet.getFirstRowNum());
	            Iterator<Cell> cellIterator = row.cellIterator();
	            while (cellIterator.hasNext()) {
	                Cell cell = cellIterator.next();
	                int columnIndex = cell.getColumnIndex();
	                sheet.autoSizeColumn(columnIndex);
	            }
	        }
	    }
	}
	
	public String handleNegativeBalance(String line) {
		StringBuffer newLine = new StringBuffer();
		int i = 0;
		for(; i < line.length() - 2; i++) {
			char c = line.charAt(i);
			char c1 = line.charAt(i+1);
			char c2 = line.charAt(i+2);
			if(c != '-') {
				newLine.append(c);
			} else if((c == '-' && c1 == '-') || (c == '-' && c1 == ' ' && c2 == '-') || (c == '-' && Character.isDigit(c1)) || (c == '-' && c1 == ' ' && Character.isDigit(c2))) {
				newLine.append("`");
				i++;
				break;
			} else {
				newLine.append(c);
			}
		}
		for(; i < line.length(); i++) {
			newLine.append(line.charAt(i));
		}
		return newLine.toString();
	}
}