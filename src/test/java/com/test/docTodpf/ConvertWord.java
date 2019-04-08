/*
 * 文件名：ConvertWord.java
 * 版权：Copyright by www.chinauip.com
 * 描述：
 * 修改人：Administrator
 * 修改时间：2017年10月11日
 * 跟踪单号：
 * 修改单号：
 * 修改内容：
 */

package com.test.docTodpf;


import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.logging.Level;
import java.util.logging.Logger;

import org.apache.poi.xwpf.converter.pdf.PdfConverter;
import org.apache.poi.xwpf.converter.pdf.PdfOptions;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
/**
 * 
 * 使用 apache poi word 转 pdf的 生成的pdf的样式与原word文档的样式相差较大
 * 〈功能详细描述〉
 * @author Administrator
 * @version 2017年10月12日
 * @see ConvertWord
 * @since
 */
public class ConvertWord {
    private static final String docName = "TestDocx.docx";
    private static final String outputlFolderPath = "d:/";


    String htmlNamePath = "docHtml.html";
    String zipName="_tmp.zip";
    File docFile = new File(outputlFolderPath+docName);
    File zipFile = new File(zipName);



      public static void main(String[] args) {
         ConvertWord cwoWord=new ConvertWord();
         cwoWord.ConvertToPDF("E:/市局电子化/旧商事系统接口文档.docx","F:/text.pdf");
         System.out.println();
    }


      
      public  void ConvertToPDF(String docPath, String pdfPath) {
          try {
              InputStream doc = new FileInputStream(new File(docPath));
              XWPFDocument document = new XWPFDocument(doc);
              PdfOptions options = PdfOptions.create();
              OutputStream out = new FileOutputStream(new File(pdfPath));
              PdfConverter.getInstance().convert(document, out, options);
          } catch (FileNotFoundException ex) {
            //  Logger.getLogger(Convert.class.getName()).log(Level.SEVERE, null, ex);
          } catch (IOException ex) {
            //  Logger.getLogger(Convert.class.getName()).log(Level.SEVERE, null, ex);
          }
      }

      public String htmlPath(){
        // d:/docHtml.html
          return outputlFolderPath+htmlNamePath;
      }

      public String zipPath(){
          // d:/_tmp.zip
          return outputlFolderPath+zipName;
      }

}