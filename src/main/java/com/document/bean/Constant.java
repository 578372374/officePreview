package com.document.bean;

import java.util.HashMap;
import java.util.Map;

public class Constant {
	//web应用名称 InitServletContextListener中进行初始化
	public static String APPLICATION_NAME = "";
	//web应用/的绝对路径 InitServletContextListener中进行初始化
	public static String REAL_PATH = "";
	//文档存放的相对位置 ->webRoot目录/docs
	public static String DOC_POSITION = "docs";
	//预览文档存放的位置 ->webRoot目录/docPreviewOnline
	public static String DOC_PREVIEW_POSITION = "docPreviewOnline";
	//文档数学公式转换成图片后，图片存放路径 ->webRoot目录/docFormulaImage
	public static String DOC_FORMULA_IMAGE_POSITION = "docFormulaImage";
	//文档转换成pdf后存放目录 ->webRoot目录/docConvertToPDF
	public static String DOC_CONVERT_TO_PDF = "docConvertToPDF";
	//JodConverterContext 对象 用于实现office文档转换为pdf InitServletContextListener中初始化
	public static JodConverterContext jodConvertercontext = null;
	
	//本地文件后缀类型与后缀名称对应关系
	public static final Map<Integer,String> LOCAL_SUFFIX_NAME_MAP = new HashMap<Integer,String>();
	static{
		LOCAL_SUFFIX_NAME_MAP.put(1,".doc");
		LOCAL_SUFFIX_NAME_MAP.put(2,".docx");
		LOCAL_SUFFIX_NAME_MAP.put(3,".ppt");
		LOCAL_SUFFIX_NAME_MAP.put(4,".pptx");
		LOCAL_SUFFIX_NAME_MAP.put(5,".xls");
		LOCAL_SUFFIX_NAME_MAP.put(6,".xlsx");
		LOCAL_SUFFIX_NAME_MAP.put(7,".pdf");
	}
}
