package com.formula.xlsx;

import java.awt.image.BufferedImage;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.StringReader;
import java.util.HashMap;
import java.util.Map;

import javax.imageio.ImageIO;
import javax.xml.transform.Source;
import javax.xml.transform.stream.StreamSource;

import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.dom4j.Document;
import org.dom4j.DocumentFactory;
import org.dom4j.Element;
import org.dom4j.io.SAXReader;
import org.junit.Test;
import org.openxmlformats.schemas.drawingml.x2006.spreadsheetDrawing.CTDrawing;
import org.openxmlformats.schemas.drawingml.x2006.spreadsheetDrawing.CTOneCellAnchor;
import org.xml.sax.InputSource;

import com.formula.util.FormulaUtil;

/**
 * xlsx公式转换成图片并写入原文件
 *
 */
public class FormulaToPng {
	/**
	 * 
	* @Description 将xlsx所有公式转换为图片并写回文件
	* @param destFilePath
	* @param pngDir
	 */
	public void dealAllXlsxMathFormula(String destFilePath,String pngDir) throws Exception{
		File file = new File(destFilePath);
		FileInputStream fis = new FileInputStream(file.getAbsolutePath());
		XSSFWorkbook xssfWorkbook = new XSSFWorkbook(fis);
		FileOutputStream fos = new FileOutputStream(file.getAbsolutePath());
		try{
			// 遍历sheet
			for (int i = 0; i < xssfWorkbook.getNumberOfSheets(); i++) {
				XSSFSheet sheet = xssfWorkbook.getSheetAt(i);
				XSSFDrawing dr = sheet.getDrawingPatriarch();
				CTDrawing drawing = dr.getCTDrawing();
				CTOneCellAnchor[] oneCells = drawing.getOneCellAnchorArray(); // 所有的公式等元素
				for (CTOneCellAnchor cellAnchor : oneCells) {
					String xml = cellAnchor.xmlText(); // 得到xml串
					// dom4j解析器的初始化
					SAXReader reader = new SAXReader(new DocumentFactory());
					Map<String, String> map = new HashMap<String, String>();
					map.put("xdr","http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing");
					map.put("m","http://schemas.openxmlformats.org/officeDocument/2006/math");
					reader.getDocumentFactory().setXPathNamespaceURIs(map); // xml文档的namespace设置

					InputSource source = new InputSource(new StringReader(xml));
					source.setEncoding("UTF-8");
					Document doc = reader.read(source);
					Element root = doc.getRootElement();
					Element e = (Element) root.selectSingleNode("//m:oMathPara"); // 用xpath得到OMML节点
					String omml = e.asXML();
					System.out.println(omml);
					ByteArrayInputStream byteArrayInputStream = new ByteArrayInputStream(omml
							.getBytes("UTF-16"));
					Source streamSource = new StreamSource(byteArrayInputStream);
					
					// omml转换成png
					String imagePath = FormulaUtil.convertOmmlToPng(streamSource, pngDir);
					// 添加图片到原公式位置
					ByteArrayOutputStream byteArrayOut = new ByteArrayOutputStream();
					BufferedImage bufferImg = ImageIO.read(new File(imagePath));
					ImageIO.write(bufferImg, "png", byteArrayOut);

					int jpegIdx = xssfWorkbook.addPicture(byteArrayOut
							.toByteArray(), XSSFWorkbook.PICTURE_TYPE_PNG);
					BufferedImage bi = ImageIO.read(new File(imagePath));
					int width = bi.getWidth();
					int height = bi.getHeight();
					int fromCol = cellAnchor.getFrom().getCol();
					int fromRow = cellAnchor.getFrom().getRow();
					XSSFClientAnchor anchor = new XSSFClientAnchor(0, 0, 0, 0,
							fromCol, fromRow, fromCol + 1 +width / 60, fromRow + 2
									+ height / 30);
					XSSFDrawing drawing2 = sheet.createDrawingPatriarch();
					drawing2.createPicture(anchor, jpegIdx);
					// 删除原有公式节点
					cellAnchor.getDomNode().getParentNode().removeChild(cellAnchor.getDomNode());
				}
			}
			//写回文件
			xssfWorkbook.write(fos);	
		}finally{
			fos.close();
			xssfWorkbook.close();
			fis.close();
		}
	}
	
	@Test
	public void test() throws Exception {
		dealAllXlsxMathFormula("D:\\workspace-sts-3.9.2.RELEASE\\officePreview\\src\\test\\resources\\formula.xlsx",".");
	}
}
