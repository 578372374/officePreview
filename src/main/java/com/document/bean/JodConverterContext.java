package com.document.bean;

import java.io.File;
import java.io.InputStream;
import java.util.Properties;

import javax.servlet.ServletContext;

import org.artofsolving.jodconverter.OfficeDocumentConverter;
import org.artofsolving.jodconverter.office.DefaultOfficeManagerConfiguration;
import org.artofsolving.jodconverter.office.OfficeManager;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

/**
 * JodConverterContext用于将office文件转换成pdf格式
 */
public class JodConverterContext {
	private Logger log = LoggerFactory.getLogger(JodConverterContext.class);
	private final OfficeManager officeManager;
	private final OfficeDocumentConverter documentConverter;

	public JodConverterContext(ServletContext context) {
		DefaultOfficeManagerConfiguration configuration = new DefaultOfficeManagerConfiguration();
		String openOfficeHomeDir = getOpenOfficeHomeDir(context);
		if (openOfficeHomeDir != null) {
		    configuration.setOfficeHome(new File(openOfficeHomeDir));
		}
		officeManager = configuration.buildOfficeManager();
		documentConverter = new OfficeDocumentConverter(officeManager);
	}

	public static void init(ServletContext context) {
		JodConverterContext jodConvertercontext = new JodConverterContext(context);
		jodConvertercontext.officeManager.start();
		Constant.jodConvertercontext = jodConvertercontext;
	}

	public static void destroy() {
		JodConverterContext jodConvertercontext = Constant.jodConvertercontext;
		jodConvertercontext.officeManager.stop();
	}
	
	public OfficeDocumentConverter getDocumentConverter() {
		return documentConverter;
	}
	
	/**
	 * 
	* @Description 得到OpenOfficeHomeDir
	 */
	public String getOpenOfficeHomeDir(ServletContext context){
		String openOfficeHomeDir = null;
		InputStream resourceAsStream = context.getResourceAsStream("/config/config.properties");
		Properties pero = null;
		try {
			pero = new Properties();
			pero.load(resourceAsStream);
			if (pero.size() > 0) {
				openOfficeHomeDir = pero.getProperty("openOffice_home_dir");
			}
		}catch(Exception e){
			log.error("初始化OpenOffice Home Dir 失败",e);
		}finally{
			if(resourceAsStream!=null){
				try{
					resourceAsStream.close();	
				}catch(Exception e){
					log.error("初始化OpenOffice Home Dir关闭文件失败",e);
				}
				pero.clear();
				pero=null;
			}
		}
		return openOfficeHomeDir;
	}
}
