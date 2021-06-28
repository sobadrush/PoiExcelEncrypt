package com.ctbc.PoiExcelEncrypt;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.commons.io.IOUtils;
import org.apache.poi.POIXMLDocumentPart;
import org.apache.poi.poifs.crypt.EncryptionInfo;
import org.apache.poi.poifs.crypt.EncryptionMode;
import org.apache.poi.poifs.crypt.Encryptor;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFRelation;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.ctbc.utils.ExcelWaterRemarkUtils;

/**
 * https://kknews.cc/zh-tw/career/ab6b9yj.html
 * https://my.oschina.net/u/4303307/blog/4523410
 */
public class TestPoiExcel {

	public static void main(String[] args) throws FileNotFoundException, IOException {
		try {
			POIFSFileSystem fs = generateEncryptExcel("1234");
			fs.writeFilesystem(new FileOutputStream("MyExcel.xlsx"));
			fs.close();
		} catch (IOException e) {
			e.printStackTrace();
		} finally {
			System.out.println("============== 完成輸出加密Excel ================");
		}

//		try(OutputStream os = new FileOutputStream("MyExcel.xlsx")) {
//			gg.writeFilesystem(os);
//			gg.close();
//		} catch (Exception e) {
//			e.printStackTrace();
//		} finally {
//			System.out.println("============== 完成輸出加密Excel ================");
//		}

//		try(POIFSFileSystem fs = new POIFSFileSystem();
//			OutputStream os = new FileOutputStream("MyExcel.xlsx");
//			Workbook workbook = new XSSFWorkbook();) {
//			EncryptionInfo encryptInfo = new EncryptionInfo(EncryptionMode.agile);
//			Encryptor enc = encryptInfo.getEncryptor();
//			enc.confirmPassword("1234");
//			
//			Sheet sheet = workbook.createSheet("sheet1");
//			sheet.createRow(0).createCell(0).setCellValue("Hello World");
//			
//			// write the workbook into the encrypted OutputStream
//			OutputStream encOutputStream = enc.getDataStream(fs);
//			workbook.write(encOutputStream);
//			workbook.close();
//			encOutputStream.close(); // this is necessary before writing out the FileSystem
//			
//			// write on disk
//			fs.writeFilesystem(os);
//		} catch (IOException | GeneralSecurityException e) {
//			e.printStackTrace();
//		} finally {
//			System.out.println("============== 完成輸出加密Excel ================");
//		}

	}

	private static List<Map<String, String>> fakeData() {
		List<Map<String, String>> empList = new ArrayList<Map<String, String>>();
		for (int i = 1; i <= 50; i++) {
			Map<String, String> empData = new HashMap<>();
			empData.put("empName", "Roger_" + i);
			empData.put("empAge", String.valueOf(20 + i));
			empData.put("empPhone", String.valueOf("0912-345-67" + i));
			empData.put("empHome", String.valueOf(i + "區"));
			empList.add(empData);
		}
		return empList;
	}

	public static POIFSFileSystem generateEncryptExcel(String excelPswd) {
		POIFSFileSystem fs = new POIFSFileSystem();
		XSSFWorkbook workbook = new XSSFWorkbook();
		try {
			EncryptionInfo encryptInfo = new EncryptionInfo(EncryptionMode.agile);
			Encryptor enc = encryptInfo.getEncryptor();
			enc.confirmPassword(excelPswd);

			XSSFSheet sheet = workbook.createSheet("sheet1");
			Row row = sheet.createRow(0);
			row.createCell(0).setCellValue("姓名");
			row.createCell(1).setCellValue("年齡");
			row.createCell(2).setCellValue("電話");
			row.createCell(3).setCellValue("居住地");

			int i = 0;
			List<Map<String, String>> empData = fakeData();
			for (Map<String, String> eMap : empData) {
				row = sheet.createRow(1 + i);
				row.createCell(0).setCellValue(eMap.get("empName"));
				row.createCell(1).setCellValue(eMap.get("empAge"));
				row.createCell(2).setCellValue(eMap.get("empPhone"));
				row.createCell(3).setCellValue(eMap.get("empHome"));
				i++;
			}
//			sheet.createRow(0).createCell(0).setCellValue("Hello World");

			ExcelWaterRemarkUtils.putWaterMarkOnEverySheetUsePng(workbook, "E:/workspace_v201909_v2/PoiExcelEncrypt/ctbc-透明-旋轉.png");
//			ExcelWaterRemarkUtils.putWaterMarkOnEverySheetUseText(workbook, "我是浮水印");

			// ----------------------------------------------------------------------------
			// 使用Excel→版面配置→背景 ( 僅可於檔案中看到浮水印，但無法於預覽列印or列印中呈現 )
			// add picture data to this workbook.
//			FileInputStream is = new FileInputStream("E:/workspace_v201909_v2/PoiExcelEncrypt/ctbc-透明.png");
//			byte[] bytes = IOUtils.toByteArray(is);
//			int pictureIdx = workbook.addPicture(bytes, Workbook.PICTURE_TYPE_PNG);
//			is.close();
//
//			// add relation from sheet to the picture data
//			POIXMLDocumentPart poixmlDocumentPart = (POIXMLDocumentPart) workbook.getAllPictures().get(pictureIdx);
//			String rID = sheet.addRelation(null, XSSFRelation.IMAGES, poixmlDocumentPart).getRelationship().getId();
//			
//			// set background picture to sheet
//			sheet.getCTWorksheet().addNewPicture().setId(rID);
			// ----------------------------------------------------------------------------
			
			// write the workbook into the encrypted OutputStream
			OutputStream encOutputStream = enc.getDataStream(fs);
			workbook.write(encOutputStream);
			workbook.close();
//			encOutputStream.close(); // this is necessary before writing out the FileSystem

			return fs;
		} catch (Exception e) {
			e.printStackTrace();
		}

		return null;
	}

	// https://codeleading.com/article/67855583660/
	public static void test() throws FileNotFoundException, IOException {
		try (XSSFWorkbook workbook = new XSSFWorkbook();
				FileOutputStream out = new FileOutputStream("CreateExcelXSSFSheetBackgroundPicture.xlsx")) {

			XSSFSheet sheet = workbook.createSheet("Sheet1");

			// add picture data to this workbook.
			FileInputStream is = new FileInputStream("dummy.png");
			byte[] bytes = IOUtils.toByteArray(is);
			int pictureIdx = workbook.addPicture(bytes, Workbook.PICTURE_TYPE_PNG);
			is.close();

			// add relation from sheet to the picture data
			POIXMLDocumentPart poixmlDocumentPart = (POIXMLDocumentPart) workbook.getAllPictures().get(pictureIdx);
			String rID = sheet.addRelation(null, XSSFRelation.IMAGES, poixmlDocumentPart).getRelationship().getId();
			
			// set background picture to sheet
			sheet.getCTWorksheet().addNewPicture().setId(rID);

			workbook.write(out);

		}
	}
}
