<%@ page language="java" contentType="text/html; charset=utf-8"
    pageEncoding="utf-8"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title>点击文件预览</title>
</head>
<body>
	<a href="<%=request.getContextPath()%>/previewDocOnline.do?docNameWithoutSuffix=预览测试&docSuffix=2">预览测试.docx</a>
	<a href="<%=request.getContextPath()%>/previewDocOnline.do?docNameWithoutSuffix=预览测试&docSuffix=4">预览测试.pptx</a>
	<a href="<%=request.getContextPath()%>/previewDocOnline.do?docNameWithoutSuffix=预览测试&docSuffix=6">预览测试.xlsx</a>
</body>
</html>