package com.formula.docx;

import java.awt.image.BufferedImage;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;

import javax.imageio.ImageIO;
import javax.xml.transform.dom.DOMSource;

import org.apache.poi.util.Units;
import org.apache.poi.xwpf.usermodel.BodyElementType;
import org.apache.poi.xwpf.usermodel.IBodyElement;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFPicture;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.junit.Test;
import org.openxmlformats.schemas.officeDocument.x2006.math.CTOMath;
import org.openxmlformats.schemas.officeDocument.x2006.math.CTOMathPara;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTP;
import org.w3c.dom.Node;

import com.formula.util.FormulaUtil;

/**
 * 
 * docx公式转换成图片并写入原文件
 *
 */
public class FormulaToPng {
	/**
	 * 
	* @Description 将docx的公式转换为图片
	* @param ctomath
	* @param pngDir
	* @param paragraph
	* @param ctp
	* @param isMathPara 如果公式节点mathPara OR math要分开处理
	 */
	public void dealDocxMathFormula(CTOMath ctomath,String pngDir,XWPFParagraph paragraph,CTP ctp,Boolean isMathPara) throws Exception{
		//将公式转换成图片
		Node node = ctomath.getDomNode();
		DOMSource source = new DOMSource(node);
		String imagePath = FormulaUtil.convertOmmlToPng(source, pngDir);
		// 添加图片
		XWPFRun run = paragraph.createRun();
		BufferedImage bi = ImageIO.read(new File(imagePath));
		int width = bi.getWidth();
		int height = bi.getHeight();
		XWPFPicture picture = run.addPicture(new FileInputStream(
				imagePath), XWPFDocument.PICTURE_TYPE_PNG, "",
				Units.toEMU(width), Units.toEMU(height));
		// 移动图片到指定位置
		Node mathNode = ctomath.getDomNode();
		Node finalMathNode = isMathPara?mathNode.getParentNode():mathNode;
		Node imageNode = picture.getCTPicture().getDomNode();
		Node finalImageNode = imageNode.getParentNode().getParentNode()
				.getParentNode().getParentNode().getParentNode();
		ctp.getDomNode().insertBefore(finalImageNode, finalMathNode);
	}
	
	/**
	 * 
	* @Description 将docx所有公式转换为图片并写回文件
	* @param destFilePath
	* @param pngDir
	 */
	public void dealAllDocxMathFormula(String destFilePath,String pngDir) throws Exception{
		FileInputStream inputStream = new FileInputStream(destFilePath);
		XWPFDocument document = new XWPFDocument(inputStream);
		FileOutputStream fos = new FileOutputStream(destFilePath);	
		try{
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

				} else if (ibodyelement.getElementType().equals(
						BodyElementType.TABLE)) {
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
		}finally{
			fos.close();
			document.close();
			inputStream.close();
		}
	}
	
	@Test
	public void test() throws Exception {
		dealAllDocxMathFormula("D:\\workspace-sts-3.9.2.RELEASE\\officePreview\\src\\test\\resources\\formula.docx",".");
	}
}
