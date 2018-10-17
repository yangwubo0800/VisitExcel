package com.example.zyj.demo;

import java.awt.FileDialog;
import java.awt.Frame;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;

import jxl.Workbook;
import jxl.format.Colour;
import jxl.format.UnderlineStyle;
import jxl.write.Label;
import jxl.write.WritableCellFormat;
import jxl.write.WritableFont;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;
import jxl.write.biff.RowsExceededException;

public class ExportExcelTest {

	public static void main(String[] args) {
		//数据
		List<HashMap<String, String>> list = new ArrayList<>();
		HashMap<String, String> map1 = new HashMap<>();
		map1.put("col1", "列1值");
		map1.put("col2", "列2值");
		map1.put("col3", "列3值");
		map1.put("col4", "列4值");
		map1.put("col5", "列5值");
		map1.put("col6", "列6值");
		map1.put("col7", "列7值");
		list.add(map1);
		HashMap<String, String> map2 = new HashMap<>();
		map2.put("col1", "列1值");
		map2.put("col2", "列2值");
		map2.put("col3", "列3值");
		map2.put("col4", "列4值");
		map2.put("col5", "列5值");
		map2.put("col6", "列6值");
		map2.put("col7", "列7值");
		list.add(map2);
		HashMap<String, String> map3 = new HashMap<>();
		map3.put("col1", "列1值");
		map3.put("col2", "列2值");
		map3.put("col3", "列3值");
		map3.put("col4", "列4值");
		map3.put("col5", "列5值");
		map3.put("col6", "列6值");
		map3.put("col7", "列7值");
		list.add(map3);
		File file = new File("testExport01.xls");
		String sheetName = "测试sheet";
		List<String> colNameList = new ArrayList<>();
		colNameList.add("字段1");
		colNameList.add("字段2");
		colNameList.add("字段3");
		colNameList.add("字段4");
		colNameList.add("字段5");
		colNameList.add("字段6");
		colNameList.add("字段7");
		
		saveToExcelStyle(list , file, sheetName, colNameList);

	}
	
	/**
	 * 保存成Excel
	 * @param list 数据
	 * @param file 文件
	 * @param sheetName sheet名
	 * @param colNameList 表头
	 * 2018-08-16 10:39
	 */
	public static void saveToExcelStyle(List<HashMap<String, String>> list,
			File file, String sheetName, List<String> colNameList) {
		OutputStream wFile;
		//表头所占行数，从0开始，如果占两行则设置为1；
		int fontRow = 1;
		//title开始行数
		int titleRow = fontRow+1;
		//数据开始行数
		int dataRow = titleRow+1;
		try {
			//wFile = new FileOutputStream(file, false);
			FileDialog fileDialog = new FileDialog(new Frame(),"选择文件夹...", FileDialog.SAVE);
//			FileAccept acceptCondition=new FileAccept("xls");
			//fileDialog.setFilenameFilter(acceptCondition);
			fileDialog.setTitle("文件另存为");
			fileDialog.setName("文件另存为");
			fileDialog.setFile(file.getName());
			fileDialog.doLayout();
			fileDialog.show(true);
			String path = fileDialog.getDirectory();
			String fileName = fileDialog.getName();
			String fileName1 = fileDialog.getFile();
			if(path==null||path==""||fileName1==null||fileName1==""){
				System.out.println("路径或文件名没有输入");
			}
			
			WritableWorkbook wbook = Workbook.createWorkbook((new File(path
					+ "\\" + fileName1)));

			WritableSheet sheet1 = wbook.createSheet(sheetName, 0);
			WritableSheet sheet2 = wbook.createSheet("第二页", 1);
			WritableSheet sheet3 = wbook.createSheet("第三页", 2);
			
			//设置表头
			WritableFont wf = new WritableFont(WritableFont.ARIAL,14,WritableFont.BOLD,false,UnderlineStyle.NO_UNDERLINE);
            WritableCellFormat wcfmt = new WritableCellFormat(wf);
            //水平居中
            wcfmt.setAlignment(jxl.format.Alignment.CENTRE);
            //垂直居中
            wcfmt.setVerticalAlignment(jxl.format.VerticalAlignment.CENTRE);
            wcfmt.setBackground(Colour.GRAY_25);
            //合并单元格（列、行）
            sheet1.mergeCells(0, 0, colNameList.size()-1, fontRow);
            Label lableCF = new Label(0,0,"测试报表",wcfmt);
            sheet1.addCell(lableCF);
           
			for (int i = 0; i < colNameList.size(); i++) {
				//title从表头的下一行开始
				sheet1.addCell(new Label(i, titleRow, colNameList.get(i)));
			}

			for (int i = 0; i < list.size(); i++) {
				for (int j = 1; j <= colNameList.size(); j++) {
					String col = list.get(i).get("col" + j);
					sheet1.addCell(new Label(j-1, dataRow+i , col));
				}

			}

			wbook.write();
			wbook.close();

		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		} catch (RowsExceededException e) {
			e.printStackTrace();
		} catch (WriteException e) {
			e.printStackTrace();
		}

	}
	
	public static void saveToExcel(List<HashMap<String, String>> list,
			File file, String sheetName, List<String> colNameList) {
		OutputStream wFile;
		try {
			//wFile = new FileOutputStream(file, false);
			FileDialog fileDialog = new FileDialog(new Frame(),"选择文件夹...", FileDialog.SAVE);
//			FileAccept acceptCondition=new FileAccept("xls");
			//fileDialog.setFilenameFilter(acceptCondition);
			fileDialog.setTitle("文件另存为");
			fileDialog.setName("文件另存为");
			fileDialog.setFile(file.getName());
			fileDialog.doLayout();
			fileDialog.show(true);
			String path = fileDialog.getDirectory();
			String fileName = fileDialog.getName();
			String fileName1 = fileDialog.getFile();
			if(path==null||path==""||fileName1==null||fileName1==""){
				System.out.println("路径或文件名没有输入");
			}
			
			WritableWorkbook wbook = Workbook.createWorkbook((new File(path
					+ "\\" + fileName1)));

			WritableSheet sheet1 = wbook.createSheet(sheetName, 0);
			WritableSheet sheet2 = wbook.createSheet("第二页", 1);
			WritableSheet sheet3 = wbook.createSheet("第三页", 2);

			for (int i = 0; i < colNameList.size(); i++) {
				sheet1.addCell(new Label(i + 1, 0, colNameList.get(i)));
			}

			for (int i = 0; i < list.size(); i++) {
				for (int j = 1; j <= colNameList.size(); j++) {
					String col = list.get(i).get("col" + j);
					sheet1.addCell(new Label(j, i + 1, col));
				}

			}

			wbook.write();
			wbook.close();

		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		} catch (RowsExceededException e) {
			e.printStackTrace();
		} catch (WriteException e) {
			e.printStackTrace();
		}

	}

}
