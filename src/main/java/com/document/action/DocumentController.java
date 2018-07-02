package com.document.action;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import org.apache.commons.io.FileUtils;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Controller;
import org.springframework.ui.ModelMap;
import org.springframework.web.bind.annotation.RequestMapping;

import com.document.bean.Constant;
import com.document.service.DocumentService;

@Controller
public class DocumentController {
	@Autowired
	DocumentService documentService;
	
	final static Logger log = LoggerFactory.getLogger(DocumentController.class);
    
	/**
	* @Description 在线预览文档
	*/
	@RequestMapping(value = "/previewDocOnline.do")
	public String previewDocOnline(String docNameWithoutSuffix,Integer docSuffix,ModelMap map,HttpServletResponse response) throws IOException{
		//全文件名
		String docNameWithSuffix = docNameWithoutSuffix + Constant.LOCAL_SUFFIX_NAME_MAP.get(docSuffix);
		//源文件位置
		String srcFilePath = Constant.REAL_PATH + Constant.DOC_POSITION + File.separator + docNameWithSuffix;
		File srcFile=new File(srcFilePath);
		//预览文件所在目录位置
		String destFilePath = Constant.REAL_PATH + Constant.DOC_PREVIEW_POSITION + File.separator + docNameWithSuffix;
		File destFile=new File(destFilePath);
		//文档中公式对应图片存放位置
		String formulaImagePosition = Constant.REAL_PATH + Constant.DOC_FORMULA_IMAGE_POSITION;
		//将源文件拷贝到预览目录
		try {
			FileUtils.copyFile(srcFile, destFile);
		} catch (Exception e) {
			log.error("在线预览拷贝文件失败", e);
			map.addAttribute("error","预览失败，拷贝文件失败！");
			return "forward:/pdfjs/web/viewerError.jsp";
		}
		// 1:doc,2:docx,3:ppt,4:pptx,5:xls,6:xlsx,7:pdf
		switch (docSuffix) {
			case 1:
			case 3:
			case 4:
				break;
			case 2:// docx
				try{
					documentService.dealAllDocxMathFormula(destFilePath, formulaImagePosition);	
				}catch(Exception e){
					log.error("在线预览，转换docx公式失败",e);
					map.addAttribute("error","预览失败，处理文档公式失败！");
					return "forward:/pdfjs/web/viewerError.jsp";
				}
				break;
			case 5:// xls
				try{
					documentService.setExcelFitToWidth(destFilePath);
				}catch(Exception e){
					log.error("在线预览，xlsx宽度自适应失败",e);
					map.addAttribute("error","预览失败，处理文档自适应宽度失败！");
					return "forward:/pdfjs/web/viewerError.jsp";
				}
				break;
			case 6:// xlsx
				try{
					documentService.dealAllXlsxMathFormula(destFilePath, formulaImagePosition);	
					documentService.setExcelFitToWidth(destFilePath);
				}catch(Exception e){
					log.error("在线预览，转换xlsx公式失败",e);
					map.addAttribute("error","预览失败，处理文档公式失败！");
					return "forward:/pdfjs/web/viewerError.jsp";
				}
				break;
			default:
				log.error("在线预览不支持的文件格式："+documentService.getLocalFileSuffixName(docSuffix));
				map.addAttribute("error","预览失败，不支持的文件格式！");
				return "forward:/pdfjs/web/viewerError.jsp";
		}
		//将原始文档转换为pdf文件
		File outputFile = new File(Constant.REAL_PATH+Constant.DOC_CONVERT_TO_PDF + File.separator + docNameWithoutSuffix + ".pdf");
		try {
			documentService.convertOfficeToPdf(destFile, outputFile);
		} catch (Exception e) {
			log.error("office文件转换为pdf失败",e);
			map.addAttribute("error","预览失败，文件转换失败！");
			return "forward:/pdfjs/web/viewerError.jsp";
		}
		//使用pdf.js预览文件
		map.addAttribute("fileToBeOpened", outputFile.getAbsolutePath().replace("\\","/"));
		return "forward:/pdfjs/web/viewer.jsp";
	}
	
	/**
	 * 
	* @Description 将文件用response输出
	 */
	@RequestMapping(value = "/pdfStreamHandeler.do")
    public void pdfStreamHandeler(String filePath, HttpServletRequest request, HttpServletResponse response) {
        File file = new File(filePath);
        byte[] data = null;
        try {
            FileInputStream input = new FileInputStream(file);
            data = new byte[input.available()];
            input.read(data);
            response.getOutputStream().write(data);
            input.close();
        } catch (Exception e) {
        	log.error("在线预览文档失败，未获取到pdf流信息",e);
        }
    }
}

