package com.document.service;

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
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamSource;

import org.apache.poi.ss.usermodel.PrintSetup;
import org.apache.poi.util.Units;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.usermodel.BodyElementType;
import org.apache.poi.xwpf.usermodel.IBodyElement;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFPicture;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.artofsolving.jodconverter.OfficeDocumentConverter;
import org.artofsolving.jodconverter.document.DocumentFormat;
import org.artofsolving.jodconverter.document.DocumentFormatRegistry;
import org.dom4j.Document;
import org.dom4j.DocumentFactory;
import org.dom4j.Element;
import org.dom4j.io.SAXReader;
import org.openxmlformats.schemas.drawingml.x2006.spreadsheetDrawing.CTDrawing;
import org.openxmlformats.schemas.drawingml.x2006.spreadsheetDrawing.CTOneCellAnchor;
import org.openxmlformats.schemas.officeDocument.x2006.math.CTOMath;
import org.openxmlformats.schemas.officeDocument.x2006.math.CTOMathPara;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTP;
import org.springframework.stereotype.Service;
import org.w3c.dom.Node;
import org.xml.sax.InputSource;

import com.document.bean.Constant;
import com.document.util.FormulaUtil;

@Service
public class DocumentService {
	
	/**
	 * 
	 * @Description 将docx的公式转换为图片
	 * @param ctomath
	 * @param pngDir
	 * @param paragraph
	 * @param ctp
	 * @param isMathPara
	 *            如果公式节点mathPara OR math要分开处理
	 */
	public void dealDocxMathFormula(CTOMath ctomath, String pngDir, XWPFParagraph paragraph, CTP ctp,
			Boolean isMathPara) throws Exception {
		// 将公式转换成图片
		Node node = ctomath.getDomNode();
		DOMSource source = new DOMSource(node);
		String imagePath = FormulaUtil.convertOmmlToPng(source, pngDir);
		// 添加图片
		XWPFRun run = paragraph.createRun();
		BufferedImage bi = ImageIO.read(new File(imagePath));
		int width = bi.getWidth();
		int height = bi.getHeight();
		XWPFPicture picture = run.addPicture(new FileInputStream(imagePath), XWPFDocument.PICTURE_TYPE_PNG, "",
				Units.toEMU(width), Units.toEMU(height));
		// 移动图片到指定位置
		Node mathNode = ctomath.getDomNode();
		Node finalMathNode = isMathPara ? mathNode.getParentNode() : mathNode;
		Node imageNode = picture.getCTPicture().getDomNode();
		Node finalImageNode = imageNode.getParentNode().getParentNode().getParentNode().getParentNode().getParentNode();
		ctp.getDomNode().insertBefore(finalImageNode, finalMathNode);
	}

	/**
	 * 
	 * @Description 将docx所有公式转换为图片并写回文件
	 * @param destFilePath
	 * @param pngDir
	 */
	public void dealAllDocxMathFormula(String destFilePath, String pngDir) throws Exception {
		FileInputStream inputStream = new FileInputStream(destFilePath);
		XWPFDocument document = new XWPFDocument(inputStream);
		FileOutputStream fos = new FileOutputStream(destFilePath);
		try {
			for (IBodyElement ibodyelement : document.getBodyElements()) {
				if (ibodyelement.getElementType().equals(BodyElementType.PARAGRAPH)) {
					XWPFParagraph paragraph = (XWPFParagraph) ibodyelement;
					CTP ctp = paragraph.getCTP();
					for (CTOMath ctomath : ctp.getOMathList()) {
						dealDocxMathFormula(ctomath, pngDir, paragraph, ctp, false);
					}

					for (CTOMathPara ctomathpara : ctp.getOMathParaList()) {
						for (CTOMath ctomath : ctomathpara.getOMathList()) {
							dealDocxMathFormula(ctomath, pngDir, paragraph, ctp, true);
						}
					}

				} else if (ibodyelement.getElementType().equals(BodyElementType.TABLE)) {
					XWPFTable table = (XWPFTable) ibodyelement;
					for (XWPFTableRow row : table.getRows()) {
						for (XWPFTableCell cell : row.getTableCells()) {
							for (XWPFParagraph paragraph : cell.getParagraphs()) {
								CTP ctp = paragraph.getCTP();
								for (CTOMath ctomath : ctp.getOMathList()) {
									dealDocxMathFormula(ctomath, pngDir, paragraph, ctp, false);
								}
								for (CTOMathPara ctomathpara : ctp.getOMathParaList()) {
									for (CTOMath ctomath : ctomathpara.getOMathList()) {
										dealDocxMathFormula(ctomath, pngDir, paragraph, ctp, true);
									}
								}
							}
						}
					}
				}
			}
			// 写回文件
			document.write(fos);
		} finally {
			fos.close();
			document.close();
			inputStream.close();
		}
	}

	/**
	 * 
	 * @Description 将xlsx所有公式转换为图片并写回文件
	 * @param destFilePath
	 * @param pngDir
	 */
	public void dealAllXlsxMathFormula(String destFilePath, String pngDir) throws Exception {
		File file = new File(destFilePath);
		FileInputStream fis = new FileInputStream(file.getAbsolutePath());
		XSSFWorkbook xssfWorkbook = new XSSFWorkbook(fis);
		FileOutputStream fos = new FileOutputStream(file.getAbsolutePath());
		try {
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
					map.put("xdr", "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing");
					map.put("m", "http://schemas.openxmlformats.org/officeDocument/2006/math");
					reader.getDocumentFactory().setXPathNamespaceURIs(map); // xml文档的namespace设置

					InputSource source = new InputSource(new StringReader(xml));
					source.setEncoding("UTF-8");
					Document doc = reader.read(source);
					Element root = doc.getRootElement();
					Element e = (Element) root.selectSingleNode("//m:oMathPara"); // 用xpath得到OMML节点
					String omml = e.asXML();
					ByteArrayInputStream byteArrayInputStream = new ByteArrayInputStream(omml.getBytes("UTF-16"));
					Source streamSource = new StreamSource(byteArrayInputStream);

					// omml转换成png
					String imagePath = FormulaUtil.convertOmmlToPng(streamSource, pngDir);
					// 添加图片到原公式位置
					ByteArrayOutputStream byteArrayOut = new ByteArrayOutputStream();
					BufferedImage bufferImg = ImageIO.read(new File(imagePath));
					ImageIO.write(bufferImg, "png", byteArrayOut);

					int jpegIdx = xssfWorkbook.addPicture(byteArrayOut.toByteArray(), XSSFWorkbook.PICTURE_TYPE_PNG);
					BufferedImage bi = ImageIO.read(new File(imagePath));
					int width = bi.getWidth();
					int height = bi.getHeight();
					int fromCol = cellAnchor.getFrom().getCol();
					int fromRow = cellAnchor.getFrom().getRow();
					XSSFClientAnchor anchor = new XSSFClientAnchor(0, 0, 0, 0, fromCol, fromRow,
							fromCol + 1 + width / 60, fromRow + 2 + height / 30);
					XSSFDrawing drawing2 = sheet.createDrawingPatriarch();
					drawing2.createPicture(anchor, jpegIdx);
					// 删除原有公式节点
					cellAnchor.getDomNode().getParentNode().removeChild(cellAnchor.getDomNode());
				}
			}
			// 写回文件
			xssfWorkbook.write(fos);
		} finally {
			fos.close();
			xssfWorkbook.close();
			fis.close();
		}
	}

	/**
	 * 设置excel宽度，保证转pdf后，可以显示在一页上
	 */
	public void setExcelFitToWidth(String destFilePath) throws Exception {
		File file = new File(destFilePath);
		FileInputStream fis = new FileInputStream(file.getAbsolutePath());
		XSSFWorkbook xssfWorkbook = new XSSFWorkbook(fis);
		FileOutputStream fos = new FileOutputStream(file.getAbsolutePath());
		try {
			// 遍历sheet 设置sheet宽度自适应，确保可以打印在一页上
			for (int i = 0; i < xssfWorkbook.getNumberOfSheets(); i++) {
				XSSFSheet sheet = xssfWorkbook.getSheetAt(i);
				sheet.setPrintGridlines(true);
				sheet.setFitToPage(true);
				PrintSetup ps = sheet.getPrintSetup();
				ps.setFitWidth((short) 1);
				ps.setFitHeight((short) 0);
			}
			// 写回文件
			xssfWorkbook.write(fos);
		} finally {
			fos.close();
			xssfWorkbook.close();
			fis.close();
		}
	}
	
	/**
	 * 根据后缀标识取真实后缀
	 * @param localSuffixType
	 * @return
	 */
	public String getLocalFileSuffixName(int localSuffixType) {
		return Constant.LOCAL_SUFFIX_NAME_MAP.get(localSuffixType);
	}
	
	/**
	 * 
	* @Description 将输入的office文件转换为pdf
	* @param inputFile 如：C:\test\input.docx
	* @param outputFile 如：C:\test\output.pdf
	 */
	public void convertOfficeToPdf(File inputFile,File outputFile) throws Exception{
		OfficeDocumentConverter converter = Constant.jodConvertercontext.getDocumentConverter();
        DocumentFormatRegistry formatRegistry = converter.getFormatRegistry();
        DocumentFormat outputFormat = formatRegistry.getFormatByExtension("pdf");
        converter.convert(inputFile, outputFile, outputFormat);	
	}
}
