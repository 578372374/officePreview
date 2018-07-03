# officePreview
#### 主要功能：
1. 实现了Office文档在线预览，禁止打印/下载
2. 支持Linux环境
3. 支持2003 2016版本Office文档，包含公式
#### 思路：
将Office文档使用JodConverter转换成PDF格式，使用PDF.js实现在线预览
#### 使用方式：
直接启动项目，访问index.jsp，即可看到在线预览功能
#### 其他
在src/test/java目录下，将项目中用到的一些核心技术做了分离，可以直接拿来单独使用
* com.formula.docx docx公式转换图片
* com.formula.xlsx xlsx公式转换图片
* com.pdf.convert Office文档转换成PDF
* com.watermark.pdf PDF文档加水印

更详细的说明可以参看与项目对应的[简书上的文章](https://www.jianshu.com/nb/27018936)
