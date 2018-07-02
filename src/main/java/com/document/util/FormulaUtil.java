package com.document.util;

import java.io.ByteArrayInputStream;
import java.io.File;
import java.io.StringWriter;

import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.transform.OutputKeys;
import javax.xml.transform.Result;
import javax.xml.transform.Source;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.stream.StreamResult;
import javax.xml.transform.stream.StreamSource;

import org.w3c.dom.Node;

import com.document.bean.Constant;

import net.sourceforge.jeuclid.MutableLayoutContext;
import net.sourceforge.jeuclid.context.LayoutContextImpl;
import net.sourceforge.jeuclid.context.Parameter;
import net.sourceforge.jeuclid.converter.Converter;

public class FormulaUtil {
	/**
	 * 
	* @Description 将数学公式转换为图片
	* @param source 
	* @param pngDir 将图片放在那个目录 比如：C:\myPicture
	 */
	public static String convertOmmlToPng(Source source,String pngDir) throws Exception {
		//将omml转换成mml
		File stylesheet = new File(Constant.REAL_PATH+"WEB-INF"+File.separator+"classes"+File.separator+"OMML2MML.XSL");
		StreamSource stylesource = new StreamSource(stylesheet);
		TransformerFactory tFac = TransformerFactory.newInstance();
		StringWriter writer = new StringWriter();
		Transformer t = tFac.newTransformer(stylesource);
		t.setOutputProperty("omit-xml-declaration", "true");
		t.setOutputProperty("indent", "yes");
		t.setOutputProperty(OutputKeys.ENCODING, "UTF-16");
		Result result = new StreamResult(writer);
		t.transform(source, result);
		String mathML = writer.toString();
		mathML = mathML
				.replaceAll(
						"xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\"",
						"");
		mathML = mathML.replaceAll("xmlns:mml", "xmlns");
		mathML = mathML.replaceAll("mml:", "");
		mathML = mathML.replaceAll(
				"xmlns=\"http://www.w3.org/1998/Math/MathML\"", "");
		mathML = mathML.replace("<?xml version=\"1.0\" encoding=\"UTF-16\"?>",
				"");
		mathML = "<math xmlns=\"http://www.w3.org/1998/Math/MathML\">\n"
				+ mathML + "</math>";
		//将mml转换成png
		Node mmlNode = DocumentBuilderFactory.newInstance()
				.newDocumentBuilder().parse(
						new ByteArrayInputStream(mathML
								.getBytes("UTF-16")))
				.getDocumentElement();
		String imagePath = pngDir+File.separator+"formula.png";
		MutableLayoutContext mlc = new LayoutContextImpl(
				LayoutContextImpl.getDefaultLayoutContext());
		//mlc.setParameter(Parameter.MATHSIZE, 8);
		// 默认0 越小越清晰	 比如清晰度排序-1 》0 》1
		mlc.setParameter(Parameter.SCRIPTLEVEL, -1);
		Converter.getInstance().convert(mmlNode,
				new File(imagePath), "image/png", mlc);
		return imagePath;
	}
}
