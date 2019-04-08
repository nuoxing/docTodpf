/*
 * 文件名：ConverterByDocx4j.java
 * 版权：Copyright by www.chinauip.com
 * 描述：
 * 修改人：Administrator
 * 修改时间：2017年10月12日
 * 跟踪单号：
 * 修改单号：
 * 修改内容：
 */

package com.test.docTodpf;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.List;

import org.docx4j.convert.out.pdf.viaXSLFO.PdfSettings;
import org.docx4j.fonts.IdentityPlusMapper;
import org.docx4j.fonts.Mapper;
import org.docx4j.fonts.PhysicalFont;
import org.docx4j.fonts.PhysicalFonts;
import org.docx4j.model.structure.MarginsWellKnown;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
/**
 * 
 * 样式还不错，就是效率比较低
 * 〈功能详细描述〉
 * @author Administrator
 * @version 2017年10月12日
 * @see ConverterByDocx4j
 * @since
 */
public class ConverterByDocx4j
{

    public ConverterByDocx4j()
    {
        // TODO Auto-generated constructor stub
    }
    
    public static void main(String[] args) {
        try {

            long start = System.currentTimeMillis();

            InputStream is = new FileInputStream(
                    new File("E:/执法/执法系统批量吊销需求.docx"));
            WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage
                    .load(is);
            List sections = wordMLPackage.getDocumentModel().getSections();
            for (int i = 0; i < sections.size(); i++) {

                System.out.println("sections Size" + sections.size());
                wordMLPackage.getDocumentModel().getSections().get(i)
                        .getPageDimensions().setMargins(MarginsWellKnown.NORMAL);
            }
            Mapper fontMapper = new IdentityPlusMapper();  
            fontMapper.put("隶书", PhysicalFonts.get("LiSu"));
            fontMapper.put("宋体",PhysicalFonts.get("SimSun"));
            fontMapper.put("微软雅黑",PhysicalFonts.get("Microsoft Yahei"));
            fontMapper.put("黑体",PhysicalFonts.get("SimHei"));
            fontMapper.put("楷体",PhysicalFonts.get("KaiTi"));
            fontMapper.put("新宋体",PhysicalFonts.get("NSimSun"));
            fontMapper.put("华文行楷", PhysicalFonts.get("STXingkai"));
            fontMapper.put("华文仿宋", PhysicalFonts.get("STFangsong"));
            fontMapper.put("宋体扩展",PhysicalFonts.get("simsun-extB"));
            fontMapper.put("仿宋",PhysicalFonts.get("FangSong"));
            fontMapper.put("仿宋_GB2312",PhysicalFonts.get("FangSong_GB2312"));
            fontMapper.put("幼圆",PhysicalFonts.get("YouYuan"));
            fontMapper.put("华文宋体",PhysicalFonts.get("STSong"));
            fontMapper.put("华文中宋",PhysicalFonts.get("STZhongsong"));  
            wordMLPackage.setFontMapper(fontMapper);
            PdfSettings pdfSettings = new PdfSettings();
            org.docx4j.convert.out.pdf.PdfConversion conversion = new org.docx4j.convert.out.pdf.viaXSLFO.Conversion(
                    wordMLPackage);

            OutputStream out = new FileOutputStream(new File(
                    "F:/text.pdf"));
          
            conversion.output(out, pdfSettings);
            System.err.println("Time taken to Generate pdf "
                    + (System.currentTimeMillis() - start) + "ms");

        } catch (Exception e) {
            e.printStackTrace();
        }
    }

}
