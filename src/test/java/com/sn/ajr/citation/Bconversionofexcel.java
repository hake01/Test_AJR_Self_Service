package com.sn.ajr.citation;
import java.io.BufferedOutputStream;

import java.io.File;
import java.io.FileInputStream; 
import java.io.FileOutputStream; 
import java.io.InputStream;
import java.io.OutputStream;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Map;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFPalette;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFWorkbook; 
public class Bconversionofexcel { 
	public static void classname() {
		System.out.println("conversion");
		File currentDirFile = new File(".");
		String rootpath1 = currentDirFile.getAbsolutePath();
		String rootpath = rootpath1.replace(".", "");
		String downloadFilepath = rootpath + "DownloadingTempfiles\\";
		String catsheet = downloadFilepath + "citaion1";
		
		File fa1 = new File(catsheet + ".xls");
		if(fa1.exists()){
		String catsheet1 = downloadFilepath + "citaion2";
		String catsheet2 = downloadFilepath + "citaion3";
		String catsheet3 = downloadFilepath + "citaion4";
		String catsheet4 = downloadFilepath + "citaion5";

		String typesheet = downloadFilepath + "savedrecs";
		String typesheet1 = downloadFilepath + "savedrecs (1)";
		String typesheet2 = downloadFilepath + "savedrecs (2)";
		String typesheet3 = downloadFilepath + "savedrecs (3)";
		String typesheet4 = downloadFilepath + "savedrecs (4)";
		try {
			Bconversionofexcel fileConversionXLSToXLXS = new Bconversionofexcel();
			String xlsFilePath = catsheet + ".xls";
			fileConversionXLSToXLXS.convertXLS2XLSX(xlsFilePath);
			String xlsFilePath1 = typesheet + ".xls";
			fileConversionXLSToXLXS.convertXLS2XLSX(xlsFilePath1);
			String xlsFilePath2 = catsheet1 + ".xls";
			fileConversionXLSToXLXS.convertXLS2XLSX(xlsFilePath2);
			String xlsFilePath3 = typesheet1 + ".xls";
			fileConversionXLSToXLXS.convertXLS2XLSX(xlsFilePath3);
			String xlsFilePath4 = catsheet2 + ".xls";
			fileConversionXLSToXLXS.convertXLS2XLSX(xlsFilePath4);
			String xlsFilePath5 = typesheet2 + ".xls";
			fileConversionXLSToXLXS.convertXLS2XLSX(xlsFilePath5);
			String xlsFilePath6 = catsheet3 + ".xls";
			fileConversionXLSToXLXS.convertXLS2XLSX(xlsFilePath6);
			String xlsFilePath7 = typesheet3 + ".xls";
			fileConversionXLSToXLXS.convertXLS2XLSX(xlsFilePath7);
			String xlsFilePath8 = catsheet4 + ".xls";
			fileConversionXLSToXLXS.convertXLS2XLSX(xlsFilePath8);
			String xlsFilePath9 = typesheet4 + ".xls";
			fileConversionXLSToXLXS.convertXLS2XLSX(xlsFilePath9);
		} catch(Exception e1) {} }
	}

	public String convertXLS2XLSX(String xlsFilePath) {
		Map cellStyleMap = new HashMap();
		String xlsxFilePath = null;
		Workbook workbookIn = null;
		File xlsxFile = null;
		Workbook workbookOut = null;
		OutputStream out = null;
		String XLSX = ".xlsx";
		try {
		try {
			InputStream inputStream = new FileInputStream(xlsFilePath);
			xlsxFilePath = xlsFilePath.substring(0, xlsFilePath.lastIndexOf('.')) + XLSX;
			workbookIn = new HSSFWorkbook(inputStream);
			xlsxFile = new File(xlsxFilePath);
			if (xlsxFile.exists())
				xlsxFile.delete();
			workbookOut = new XSSFWorkbook();
			int sheetCnt = workbookIn.getNumberOfSheets();

			for (int i = 0; i < sheetCnt; i++) {
				Sheet sheetIn = workbookIn.getSheetAt(i);
				Sheet sheetOut = workbookOut.createSheet(sheetIn.getSheetName());
				Iterator rowIt = sheetIn.rowIterator();
				while (rowIt.hasNext()) {
					Row rowIn = (Row) rowIt.next();
					Row rowOut = sheetOut.createRow(rowIn.getRowNum());
					copyRowProperties(rowOut, rowIn,cellStyleMap);
				}
			}
			out = new BufferedOutputStream(new FileOutputStream(xlsxFile));
			workbookOut.write(out);
		} catch (Exception ex) {
			System.err.println("Exception Occured inside transFormXLS2XLSX :: file Name :: " + xlsFilePath
					+ ":: reason ::" + ex.getMessage());
			ex.printStackTrace();
			xlsxFilePath = null;
		} finally {
			try {
				if (workbookOut != null)
					workbookOut.close();
				if (workbookIn != null)
					workbookIn.close();
				if (out != null)
					out.close();
			} catch (Exception ex) {
				ex.printStackTrace();
			}
		}
		
	} catch(Exception e1) {}
		return xlsxFilePath;
	}

	private void copyRowProperties(Row rowOut, Row rowIn, Map cellStyleMap) {
		try {
		rowOut.setRowNum(rowIn.getRowNum());
		rowOut.setHeight(rowIn.getHeight());
		rowOut.setHeightInPoints(rowIn.getHeightInPoints());
		rowOut.setZeroHeight(rowIn.getZeroHeight());
		Iterator cellIt = rowIn.cellIterator();
		while (cellIt.hasNext()) {
			Cell cellIn = (Cell) cellIt.next();
			Cell cellOut = rowOut.createCell(cellIn.getColumnIndex(), cellIn.getCellType());
			rowOut.getSheet().setColumnWidth(cellOut.getColumnIndex(),
					rowIn.getSheet().getColumnWidth(cellIn.getColumnIndex()));
			copyCellProperties(cellOut, cellIn, cellStyleMap);
		}
		} catch(Exception e1) {}
	}

	private void copyCellProperties(Cell cellOut, Cell cellIn, Map cellStyleMap) {
try {
		Workbook wbOut = cellOut.getSheet().getWorkbook();
		HSSFPalette hssfPalette = ((HSSFWorkbook) cellIn.getSheet().getWorkbook()).getCustomPalette();
		switch (cellIn.getCellType()) {
		case Cell.CELL_TYPE_BLANK:
			break;

		case Cell.CELL_TYPE_BOOLEAN:
			cellOut.setCellValue(cellIn.getBooleanCellValue());
			break;

		case Cell.CELL_TYPE_ERROR:
			cellOut.setCellValue(cellIn.getErrorCellValue());
			break;

		case Cell.CELL_TYPE_FORMULA:
			cellOut.setCellFormula(cellIn.getCellFormula());
			break;

		case Cell.CELL_TYPE_NUMERIC:
			cellOut.setCellValue(cellIn.getNumericCellValue());
			break;

		case Cell.CELL_TYPE_STRING:
			cellOut.setCellValue(cellIn.getStringCellValue());
			break;
		}
		HSSFCellStyle styleIn = (HSSFCellStyle) cellIn.getCellStyle();
		XSSFCellStyle styleOut = null;
		if (cellStyleMap.get(styleIn.getIndex()) != null) {
			styleOut = (XSSFCellStyle) cellStyleMap.get(styleIn.getIndex());
		} else {
			styleOut = (XSSFCellStyle) wbOut.createCellStyle();
			styleOut.setAlignment(styleIn.getAlignment());
			DataFormat format = wbOut.createDataFormat();
			styleOut.setDataFormat(format.getFormat(styleIn.getDataFormatString()));
			HSSFColor forgroundColor = styleIn.getFillForegroundColorColor();
			if (forgroundColor != null) {
				short[] foregroundColorValues = forgroundColor.getTriplet();
				styleOut.setFillForegroundColor(new XSSFColor(new java.awt.Color(foregroundColorValues[0],
						foregroundColorValues[1], foregroundColorValues[2])));
				styleOut.setFillPattern(styleIn.getFillPattern());
			}
			styleOut.setFillPattern(styleIn.getFillPattern());
			styleOut.setBorderBottom(styleIn.getBorderBottom());
			styleOut.setBorderLeft(styleIn.getBorderLeft());
			styleOut.setBorderRight(styleIn.getBorderRight());
			styleOut.setBorderTop(styleIn.getBorderTop());
			HSSFColor bottom = hssfPalette.getColor(styleIn.getBottomBorderColor());
			if (bottom != null) {
				short[] bottomColorArray = bottom.getTriplet();
				styleOut.setBottomBorderColor(new XSSFColor(new java.awt.Color(bottomColorArray[0],
						bottomColorArray[1], bottomColorArray[2])));
			}
			HSSFColor top = hssfPalette.getColor(styleIn.getTopBorderColor());
			if (top != null) {
				short[] topColorArray = top.getTriplet();
				styleOut.setTopBorderColor(new XSSFColor(new java.awt.Color(topColorArray[0], topColorArray[1],
						topColorArray[2])));
			}
			HSSFColor left = hssfPalette.getColor(styleIn.getLeftBorderColor());
			if (left != null) {
				short[] leftColorArray = left.getTriplet();
				styleOut.setLeftBorderColor(new XSSFColor(new java.awt.Color(leftColorArray[0], leftColorArray[1],
						leftColorArray[2])));
			}
			HSSFColor right = hssfPalette.getColor(styleIn.getRightBorderColor());
			if (right != null) {
				short[] rightColorArray = right.getTriplet();
				styleOut.setRightBorderColor(new XSSFColor(new java.awt.Color(rightColorArray[0], rightColorArray[1],
						rightColorArray[2])));
			}
			styleOut.setVerticalAlignment(styleIn.getVerticalAlignment());
			styleOut.setHidden(styleIn.getHidden());
			styleOut.setIndention(styleIn.getIndention());
			styleOut.setLocked(styleIn.getLocked());
			styleOut.setRotation(styleIn.getRotation());
			styleOut.setShrinkToFit(styleIn.getShrinkToFit());
			styleOut.setVerticalAlignment(styleIn.getVerticalAlignment());
			styleOut.setWrapText(styleIn.getWrapText());
			cellOut.setCellComment(cellIn.getCellComment());
			cellStyleMap.put(styleIn.getIndex(), styleOut);
		}
		cellOut.setCellStyle(styleOut);
		} catch(Exception e1) {}
	}


	
}
