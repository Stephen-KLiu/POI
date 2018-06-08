package com.cn.POI;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

public class MyPOI
{
	@Test
	public void write() throws IOException
	{
		//创建一个工作簿对象
		//Workbook workbook = new HSSFWorkbook();//.xls文件
		Workbook workbook = new XSSFWorkbook();//.xlsx文件
		//创建一个工作表
		Sheet sheet = workbook.createSheet("mySheet01");
		//创建指定的行，从0开始计数
		Row row = sheet.createRow(4);
		//创建指定的单元格，从0开始计数
		Cell cell = row.createCell(4);
		//设置该单元格的值
		cell.setCellValue("zhangsan");
		
		//将该工作簿写到指定位置
		//workbook.write(new FileOutputStream(new File("E:\JavaProject\Java02\Day31_POI\\write.xls")));
		workbook.write(new FileOutputStream(new File("E:\\JavaProject\\Java02\\Day31_POI\\write.xlsx")));
		workbook.close();
	}
	
	@Test
	public void read() throws FileNotFoundException, IOException
	{
		//获取指定的工作簿
		//Workbook workbook = new HSSFWorkbook(new FileInputStream(new File("E:\JavaProject\Java02\Day31_POI\\read.xlsx")));
		Workbook workbook = new XSSFWorkbook(new FileInputStream(new File("E:\\JavaProject\\Java02\\Day31_POI\\read.xlsx")));
		//获取指定的工作表
		Sheet sheet = workbook.getSheet("mySheet01");
		System.out.println(sheet.getPhysicalNumberOfRows());//获取Excel表中已被使用的行数，该行只要有一个单元内容不为空则该行就被使用了
		System.out.println(sheet.getLastRowNum());//获取Excel表中最后一个被使用的行号-1
		for (int i = 0; i <= sheet.getLastRowNum(); i++)
		{
			/**注意：若行内容为空（该行未被使用），则获取的row对象为null
			 * 若单元格内容为空，则获取的cell对象为null
			 * */
			//获取指定的行
			Row row = sheet.getRow(i);
			if (row == null)
			{
				continue;
			}
			//获取指定的单元格
			Cell nameCell = row.getCell(0);
			Cell genderCell = row.getCell(1);
			Cell passwordCell = row.getCell(2);
			System.out.print(nameCell.getStringCellValue()+"		");
			System.out.print(genderCell.getStringCellValue()+"		");
			System.out.println(passwordCell.getStringCellValue());
		}
		workbook.close();
	}
}
