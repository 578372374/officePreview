package com.pdf.convert;

import java.io.File;

import org.artofsolving.jodconverter.OfficeDocumentConverter;
import org.artofsolving.jodconverter.document.DocumentFormat;
import org.artofsolving.jodconverter.document.DocumentFormatRegistry;
import org.artofsolving.jodconverter.office.DefaultOfficeManagerConfiguration;
import org.artofsolving.jodconverter.office.OfficeManager;
import org.junit.Test;

/**
 * 使用OpenOffice将office文件转换成pdf
 */
public class JodConverterContext {
	private final OfficeManager officeManager;
	private final OfficeDocumentConverter documentConverter;

	public JodConverterContext() {
		DefaultOfficeManagerConfiguration configuration = new DefaultOfficeManagerConfiguration();
		configuration.setOfficeHome(new File("C:\\Program Files (x86)\\OpenOffice 4"));
		officeManager = configuration.buildOfficeManager();
		documentConverter = new OfficeDocumentConverter(officeManager);
	}

	public void init() {
		this.officeManager.start();
	}

	public void destroy() {
		this.officeManager.stop();
	}
	
	public OfficeDocumentConverter getDocumentConverter() {
		return documentConverter;
	}
	
	@Test
	public void test() {
		JodConverterContext jodConverterContext = new JodConverterContext();
		jodConverterContext.init();
		OfficeDocumentConverter converter = jodConverterContext.getDocumentConverter();
        DocumentFormatRegistry formatRegistry = converter.getFormatRegistry();
        DocumentFormat outputFormat = formatRegistry.getFormatByExtension("pdf");
        
        //转换docx
        File inputFile = new File("D:\\workspace-sts-3.9.2.RELEASE\\officePreview\\src\\test\\resources\\formula.docx");
        File outputFile = new File("D:\\workspace-sts-3.9.2.RELEASE\\officePreview\\src\\test\\resources\\formula_docx.pdf");
        converter.convert(inputFile,outputFile,outputFormat);	
        
        //如果是excel，需要先调整excel的宽度，否则转换成pdf不在一页纸上
        String excelPath = "D:\\workspace-sts-3.9.2.RELEASE\\officePreview\\src\\test\\resources\\formula.xlsx";
        File inputFile2 = new File(excelPath);
        File outputFile2 = new File("D:\\workspace-sts-3.9.2.RELEASE\\officePreview\\src\\test\\resources\\formula_xlsx.pdf");
        try {
			ConverterUtil.setExcelFitToWidth(excelPath);
		} catch (Exception e) {
			e.printStackTrace();
		}
        converter.convert(inputFile2,outputFile2,outputFormat);	
        
        jodConverterContext.destroy();
	}
}
