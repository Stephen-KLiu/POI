package com.cn.POI;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.hssf.util.HSSFColor.HSSFColorPredefined;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

public class MyPOIStyle
{
	@Test
	public void write() throws IOException
	{
		Workbook workbook = new XSSFWorkbook();
		Sheet sheet = workbook.createSheet("mySheet01");
		sheet.setDefaultColumnWidth(20);//设置默认列宽
		Row row = sheet.createRow(0);
		
		//单元格样式
		CellStyle cellStyle = workbook.createCellStyle();
		//1).对齐方式
		cellStyle.setAlignment(HorizontalAlignment.CENTER);//水平居中
		cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);//垂直居中
		//2).字体
		Font font = workbook.createFont();
		font.setFontName("宋体");//设置字体
		font.setFontHeightInPoints((short) 20);//字体大小
		font.setColor(HSSFColorPredefined.RED.getIndex());//设置字体颜色
		font.setBold(true);//是否加粗
		font.setItalic(false);//是否斜体
		cellStyle.setFont(font);
		//3).填充色
		cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		cellStyle.setFillForegroundColor(HSSFColorPredefined.GREY_25_PERCENT.getIndex());
		
		Cell cell = row.createCell(0);
		cell.setCellValue("用户信息列表");
		cell.setCellStyle(cellStyle);
		
		//合并单元格
		CellRangeAddress cellRangeAddress = new CellRangeAddress(0, 1, 0, 3);
		sheet.addMergedRegion(cellRangeAddress);
		
		workbook.write(new FileOutputStream(new File("E:\\JavaProject\\01.xlsx")));
		workbook.close();
	}
	
	@Test
	public void read() throws FileNotFoundException, IOException
	{
		Workbook workbook = new XSSFWorkbook(new FileInputStream(new File("E:\\JavaProject\\01.xlsx")));
		Sheet sheet = workbook.getSheet("mySheet01");
		Row row = sheet.getRow(4);
		Cell cell = row.getCell(4);
		
		System.out.println(cell.getStringCellValue());
		workbook.close();
	}
}
