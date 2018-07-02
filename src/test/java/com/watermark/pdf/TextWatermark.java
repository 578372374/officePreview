package com.watermark.pdf;
 
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

import com.itextpdf.text.BaseColor;
import com.itextpdf.text.DocumentException;
import com.itextpdf.text.Element;
import com.itextpdf.text.Rectangle;
import com.itextpdf.text.pdf.BaseFont;
import com.itextpdf.text.pdf.PdfContentByte;
import com.itextpdf.text.pdf.PdfGState;
import com.itextpdf.text.pdf.PdfReader;
import com.itextpdf.text.pdf.PdfStamper;

/**
 * 添加文字水印
 * ①支持中文
 * ②透明/非透明
 * ③是否覆盖原内容
 * ④位置灵活：侧边 顶部 正文任意位置
 */
public class TextWatermark {
 
    public static final String SRC = "D:\\hero.pdf";
    public static final String DEST = "D:\\hero_watermarked.pdf";
 
    public static void main(String[] args) throws IOException, DocumentException {
        File file = new File(DEST);
        file.getParentFile().mkdirs();
        new TextWatermark().manipulatePdf(SRC, DEST);
    }
 
    public void manipulatePdf(String src, String dest) throws IOException, DocumentException {
        PdfReader reader = new PdfReader(src);
        PdfStamper stamper = new PdfStamper(reader, new FileOutputStream(dest));
        //需要引入字体文件比如这里的新宋体，否则不支持中文
        BaseFont base = BaseFont.createFont(  
                "D:\\workspace-sts-3.9.2.RELEASE\\officePreview\\src\\main\\resources\\SIMSUN.TTC,1", BaseFont.IDENTITY_H, BaseFont.EMBEDDED); 

        //水印在内容的下层 under
        PdfContentByte under = stamper.getUnderContent(1);
        under.setFontAndSize(base, 40);
        under.setColorFill(BaseColor.GRAY);  
        String underText = "该水印在下层";
        under.showTextAligned(Element.ALIGN_CENTER, underText, 200,250, 45); 
        
        //水印在内容的上层 over
        PdfContentByte over = stamper.getOverContent(1);
        over.setFontAndSize(base, 40);
        over.setColorFill(BaseColor.GRAY);
        String overText = "该水印在上层";
        over.showTextAligned(Element.ALIGN_CENTER, overText, 200,450, 45); 
        
        //透明水印
        over.saveState();
        PdfGState gs1 = new PdfGState();
        gs1.setFillOpacity(0.5f);
        over.setGState(gs1);
        String overOpacityText = "我是透明水印";
        over.showTextAligned(Element.ALIGN_CENTER, overOpacityText, 200,650, 45); 
        over.restoreState();
        
        //根据页面大小信息，给侧边添加水印
        Rectangle pageSize;
        pageSize = reader.getPageSizeWithRotation(1);//页面大小信息
        float x = pageSize.getLeft();//页面左边坐标
        float y = (pageSize.getTop() + pageSize.getBottom()) / 2;//页面竖直方向中间坐标
        over.setFontAndSize(base, 20);
        over.setColorFill(BaseColor.GRAY);
        String sideText = "该水印在侧面";
        over.showTextAligned(Element.ALIGN_CENTER, sideText, x+18, y, 90); 
        
        //给顶部添加水印
        float x2 = pageSize.getWidth()/2;//页面水平方向中间坐标
        float y2 = pageSize.getTop(20);//页面距离顶部20的位置
        over.setFontAndSize(base, 10);
        over.setColorFill(BaseColor.GRAY);
        String topText = "该水印在顶部";
        over.showTextAligned(Element.ALIGN_CENTER, topText, x2, y2, 0); 
        
        stamper.close();
        reader.close();
    }
}