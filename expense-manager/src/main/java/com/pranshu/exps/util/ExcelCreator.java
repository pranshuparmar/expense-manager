package com.pranshu.exps.util;

import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import javax.script.ScriptEngine;
import javax.script.ScriptEngineManager;
import javax.script.ScriptException;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.RegionUtil;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelCreator {

	Workbook workbook = new XSSFWorkbook();
	Sheet sheet = workbook.createSheet();
	CellStyle cellStyleNormal = workbook.createCellStyle();
	CellStyle cellStyleBold = workbook.createCellStyle();
	Font boldFont = workbook.createFont();
	ArrayList<Double> dayTotalList = new ArrayList<Double>();
	int totRows = 0;

	public void convertRecords(List<Object> records, String outputFileName) throws ScriptException {
		Iterator<Object> itr = records.iterator();
		int rowNum = 0, colNum = 0;
		double dayTotal = 0, monthTotal = 0;

		Row monthHeader = sheet.createRow(rowNum++);
		++totRows;
		monthHeader.createCell(colNum);

		Row dateRow = sheet.createRow(rowNum++);
		++totRows;

		while (itr.hasNext()) {
			String line = itr.next().toString().trim();
			String[] lineTokens = line.split("-");

			if (lineTokens.length == 1 && !line.isEmpty()) {
				if (dayTotal != 0) {
					dayTotalList.add(dayTotal);
					monthTotal += dayTotal;
					dayTotal = 0;
					rowNum = 2;
				}
				Cell cell = dateRow.createCell(colNum);
				cell.setCellValue(line.trim());
				colNum = colNum + 2;
				sheet.addMergedRegion(
						new CellRangeAddress(dateRow.getRowNum(), dateRow.getRowNum(), colNum - 2, colNum - 1));
			} else if (lineTokens.length == 2) {
				double priceVal = 0;
				if (lineTokens[1].trim().split("-").length > 1 || lineTokens[1].trim().split("\\+").length > 1)
					priceVal = Double.valueOf(evaluateFinalAmount(lineTokens[1].trim()));
				else
					priceVal = Double.valueOf(lineTokens[1].trim());
				dayTotal += createItemPriceCells(rowNum++, colNum, lineTokens[0].trim(), priceVal);
			} else if (line.indexOf("-") != line.lastIndexOf("-")) {
				String[] newLine = handleNegativeBalance(line).split("`");
				double priceVal = 0;
				if (newLine[1].trim().split("-").length > 1 || newLine[1].trim().split("\\+").length > 1)
					priceVal = Double.valueOf(evaluateFinalAmount(newLine[1].trim()));
				else
					priceVal = Double.valueOf(newLine[1].trim());
				dayTotal += createItemPriceCells(rowNum++, colNum, newLine[0].trim(), priceVal);
			}
		}

		dayTotalList.add(dayTotal);

		int col = 0;
		sheet.createRow(totRows);
		for (Double total : dayTotalList) {
			col += 2;
			createItemPriceCells(totRows, col, "Day Total", total);
		}

		createItemPriceCells(++totRows, 2, "Month Total", monthTotal);
		sheet.addMergedRegion(new CellRangeAddress(totRows - 1, totRows - 1, 1, col - 1));

		sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, col - 1));

		applyBorder(totRows, col);

		autoSizeColumns();

		try {
			FileOutputStream fileOut = new FileOutputStream(outputFileName);
			workbook.write(fileOut);
			fileOut.close();

			// Closing the workbook
			workbook.close();
		} catch (Exception ex) {
		}
	}

	private void autoSizeColumns() {
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

	private String handleNegativeBalance(String line) {
		StringBuffer newLine = new StringBuffer();
		int i = 0;
		for (; i < line.length() - 2; i++) {
			char c = line.charAt(i);
			char c1 = line.charAt(i + 1);
			char c2 = line.charAt(i + 2);
			if (c != '-') {
				newLine.append(c);
			} else if ((c == '-' && c1 == '-') || (c == '-' && c1 == ' ' && c2 == '-')
					|| (c == '-' && Character.isDigit(c1)) || (c == '-' && c1 == ' ' && Character.isDigit(c2))) {
				newLine.append("`");
				i++;
				break;
			} else {
				newLine.append(c);
			}
		}
		for (; i < line.length(); i++) {
			newLine.append(line.charAt(i));
		}
		return newLine.toString();
	}

	private String evaluateFinalAmount(String amounts) throws ScriptException {
		ScriptEngineManager mgr = new ScriptEngineManager();
		ScriptEngine engine = mgr.getEngineByName("JavaScript");
		return engine.eval(amounts).toString();
	}

	private double createItemPriceCells(int row, int col, String itemVal, double priceVal) {
		Row dataRow = sheet.getRow(row);
		if (dataRow == null) {
			dataRow = sheet.createRow(row);
			++totRows;
		}
		Cell item = dataRow.createCell(col - 2);
		Cell price = dataRow.createCell(col - 1);
		item.setCellValue(itemVal);
		price.setCellValue(priceVal);
		return price.getNumericCellValue();
	}

	private void applyBorder(int lastRow, int lastCol) {
		cellStyleNormal.setAlignment(HorizontalAlignment.CENTER);
		cellStyleNormal.setBorderLeft(BorderStyle.THIN);
		cellStyleNormal.setBorderRight(BorderStyle.THIN);
		cellStyleNormal.setBorderBottom(BorderStyle.THIN);
		cellStyleNormal.setBorderTop(BorderStyle.THIN);

		boldFont.setBold(true);
		cellStyleBold.setAlignment(HorizontalAlignment.CENTER);
		cellStyleBold.setFont(boldFont);
		cellStyleBold.setBorderLeft(BorderStyle.THICK);
		cellStyleBold.setBorderRight(BorderStyle.THICK);
		cellStyleBold.setBorderBottom(BorderStyle.THICK);
		cellStyleBold.setBorderTop(BorderStyle.THICK);

		for (int i = 0; i < lastRow; i++) {
			Row row = sheet.getRow(i);
			for (int j = 0; j < lastCol; j++) {
				Cell cell = row.getCell(j);
				if (cell == null)
					cell = row.createCell(j);
				if (i == 0 || i == 1 || i == lastRow - 2 || i == lastRow - 1)
					cell.setCellStyle(cellStyleBold);
				else
					cell.setCellStyle(cellStyleNormal);
			}
		}

		CellRangeAddress region = new CellRangeAddress(0, lastRow - 1, 0, lastCol - 1);
		RegionUtil.setBorderBottom(BorderStyle.THICK, region, sheet);
		RegionUtil.setBorderTop(BorderStyle.THICK, region, sheet);
		RegionUtil.setBorderLeft(BorderStyle.THICK, region, sheet);
		RegionUtil.setBorderRight(BorderStyle.THICK, region, sheet);
	}
}
