package com.ctbc.utils;

import java.awt.AlphaComposite;
import java.awt.BasicStroke;
import java.awt.Color;
import java.awt.Font;
import java.awt.Graphics2D;
import java.awt.Transparency;
import java.awt.font.FontRenderContext;
import java.awt.geom.Rectangle2D;
import java.awt.image.BufferedImage;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.IOException;
import java.io.InputStream;

import javax.imageio.ImageIO;

import org.apache.commons.io.FileUtils;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.Picture;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

public class ExcelWaterRemarkUtils {
	/**
	 * 為Excel打上水印工具函數 請自行確保參數值，以保證水印圖片之間不會覆蓋。在計算水印的位置的時候，並沒有考慮到單元格合併的情況，請注意
	 *
	 * @param wb: Excel Workbook
	 *
	 * @param sheet: 需要打水印的Excel
	 *
	 * @param waterRemarkPath: 水印地址，classPath，目前只支持 png格式 的圖片，
	 * 因為非png格式的圖片打到Excel上後可能會有圖片變紅的問題，且不容易做出透明效果。
	 * 同時請注意傳入的地址格式，應該為類似："\\excelTemplate\\test.png"
	 *
	 * @param startXCol: 水印起始列
	 *
	 * @param startYRow: 水印起始行
	 *
	 * @param betweenXCol: 水印橫向之間間隔多少列
	 *
	 * @param betweenYRow: 水印縱向之間間隔多少行
	 *
	 * @param XCount: 橫向共有水印多少個
	 *
	 * @param YCount: 縱向共有水印多少個
	 *
	 * @param waterRemarkWidth: 水印圖片寬度為多少列
	 *
	 * @param waterRemarkHeight: 水印圖片高度為多少行
	 *
	 * @throws IOException:
	 *
	 * ref. https://www.jianshu.com/p/5ebf2217f0be
	 */
	public static void putWaterRemarkToExcel(
			Workbook wb, Sheet sheet, 
			byte[] waterRemarkImgByteArray, 
			int startXCol, int startYRow, 
			int betweenXCol, int betweenYRow, 
			int XCount, int YCount, 
			int waterRemarkWidth,
			int waterRemarkHeight) throws IOException {

		// 加載圖片
		try (ByteArrayOutputStream byteArrayOut = new ByteArrayOutputStream();
			 InputStream imageIn = new ByteArrayInputStream(waterRemarkImgByteArray);) 
		{
			if (null == imageIn || imageIn.available() < 1) {
				throw new RuntimeException("向Excel上面打印水印，讀取水印圖片失敗(1)");
			}
			BufferedImage bufferImg = ImageIO.read(imageIn);
			if (null == bufferImg) {
				throw new RuntimeException("向Excel上面打印水印，讀取水印圖片失敗(2)");
			}
			ImageIO.write(bufferImg, "png", byteArrayOut);
			
			// 開始打水印
			Drawing drawing = sheet.createDrawingPatriarch();

			// 按照共需打印多少行水印進行循環
			for (int yCount = 0; yCount < YCount; yCount++) {
				// 按照每行需要打印多少個水印進行循環
				for (int xCount = 0; xCount < XCount; xCount++) {
					// 創建水印圖片位置
					int xIndexInteger = startXCol + (xCount * waterRemarkWidth) + (xCount * betweenXCol);
					int yIndexInteger = startYRow + (yCount * waterRemarkHeight) + (yCount * betweenYRow);
					/*
					* 參數定義： 第一個參數是（x軸的開始節點）； 第二個參數是（是y軸的開始節點）； 第三個參數是（是x軸的結束節點）； 第四個參數是（是y軸的結束節點）；
					* 第五個參數是（是從Excel的第幾列開始插入圖片，從0開始計數）； 第六個參數是（是從excel的第幾行開始插入圖片，從0開始計數）；
					* 第七個參數是（圖片寬度，共多少列）； 第8個參數是（圖片高度，共多少行）；
					*/
					ClientAnchor anchor = drawing.createAnchor(0, 0, 0, 0, xIndexInteger, yIndexInteger, xIndexInteger + waterRemarkWidth, yIndexInteger + waterRemarkHeight);

					Picture pic = drawing.createPicture(anchor, wb.addPicture(byteArrayOut.toByteArray(), Workbook.PICTURE_TYPE_PNG));
					pic.resize();
				}
			}
		}

	}

	/**
	 * @param content : 要轉成水印的文字
	 * @param path    : 水印圖片路徑
	 * @throws IOException
	 */
	public static byte[] createWaterMark(String content) throws IOException {
		Integer width = 300;
		Integer height = 200;
		BufferedImage image = new BufferedImage(width, height, BufferedImage.TYPE_INT_RGB); // 獲取bufferedImage對象
		String fontType = "Microsoft Jhenghei";
		Integer fontStyle = Font.PLAIN;
		Integer fontSize = 50;
		Font font = new Font(fontType, fontStyle, fontSize);
		Graphics2D g2d = image.createGraphics(); // 獲取Graphics2d對象
		image = g2d.getDeviceConfiguration().createCompatibleImage(width, height, Transparency.TRANSLUCENT);
		g2d.dispose();
		g2d = image.createGraphics();
		g2d.setColor(new Color(0, 0, 0, 80)); // 設置字體顏色和透明度
		g2d.setStroke(new BasicStroke(1)); // 設置字體
		g2d.setFont(font); // 設置字體類型 加粗 大小
		g2d.rotate(Math.toRadians(-10), (double) image.getWidth() / 2, (double) image.getHeight() / 2); // 設置傾斜度
		FontRenderContext context = g2d.getFontRenderContext();
		Rectangle2D bounds = font.getStringBounds(content, context);
		double x = (width - bounds.getWidth()) / 2;
		double y = (height - bounds.getHeight()) / 2;
		double ascent = -bounds.getY();
		double baseY = y + ascent;
		// 寫入水印文字原定高度過小，所以累計寫水印，增加高度
		g2d.drawString(content, (int) x, (int) baseY);
		// 設置透明度
		g2d.setComposite(AlphaComposite.getInstance(AlphaComposite.SRC_OVER));
		// 釋放對象
		g2d.dispose();
//		ImageIO.write(image, "png", new File(path));
		return toByteArray(image, "png");
	}

    // convert BufferedImage to byte[]
    public static byte[] toByteArray(BufferedImage bi, String format) throws IOException {
        ByteArrayOutputStream baos = new ByteArrayOutputStream();
        ImageIO.write(bi, format, baos);
        byte[] bytes = baos.toByteArray();
        return bytes;
    }
	
	public static void putWaterMarkOnEverySheetUsePng(Workbook wb, String imgPath) throws IOException {
		
		// 校驗傳入的水印圖片格式
		if (!imgPath.endsWith("png") && !imgPath.endsWith("PNG")) {
			throw new RuntimeException("僅支援png格式圖片");
		}
		
		// 獲取excel sheet個數
		int sheets = wb.getNumberOfSheets();
		// 循環sheet給每個sheet添加水印
		for (int i = 0; i < sheets; i++) {
			Sheet sheet = wb.getSheetAt(i);
			// excel加密只讀
			// sheet.protectSheet(UUID.randomUUID().toString());
			// 獲取excel實際所佔行
			int row = sheet.getFirstRowNum() + sheet.getLastRowNum();
			// 獲取excel實際所佔列
			int cell = sheet.getRow(sheet.getFirstRowNum()).getLastCellNum() + 1;
			// 根據行與列計算實際所需多少水印
			ExcelWaterRemarkUtils.putWaterRemarkToExcel(wb, sheet, FileUtils.readFileToByteArray(new File(imgPath)), 0, 0, 10, 10, cell / 5 + 1, row / 5 + 1, 0, 0);
		}
	}

	public static void putWaterMarkOnEverySheetUseText(Workbook wb, String waterMarkText) throws IOException {
		byte[] waterRemarkImgByteArray = createWaterMark(waterMarkText); // 將傳入的text轉成浮水印圖片
		// 獲取excel sheet個數
		int sheets = wb.getNumberOfSheets();
		// 循環sheet給每個sheet添加水印
		for (int i = 0; i < sheets; i++) {
			Sheet sheet = wb.getSheetAt(i);
			// excel加密只讀
			// sheet.protectSheet(UUID.randomUUID().toString());
			// 獲取excel實際所佔行
			int row = sheet.getFirstRowNum() + sheet.getLastRowNum();
			// 獲取excel實際所佔列
			int cell = sheet.getRow(sheet.getFirstRowNum()).getLastCellNum() + 1;
			// 根據行與列計算實際所需多少水印
			ExcelWaterRemarkUtils.putWaterRemarkToExcel(wb, sheet, waterRemarkImgByteArray, 0, 0, 5, 5, cell / 5 + 1, row / 5 + 1, 0, 0);
		}
	}
}
