package com.tfd.poi2html;

import com.tfd.POIUtils;
import com.tfd.pdf2html.PDF2Image;
import org.apache.poi.hslf.blip.PNG;
import top.lllyl2012.html.utils.File2HtmlUtil;
import top.lllyl2012.html.utils.File2byte;
import top.lllyl2012.html.utils.LogicUtil;

import javax.servlet.http.HttpServletResponse;
import java.io.File;
import java.io.IOException;
import java.io.OutputStream;

/**
 * @author TangFD@HF 2018/5/29
 */
public class POI2Html {
    public static void main(String[] args) throws Exception {
//        poi2Html("xls/test.xls", "xls", "xls/img", "img");
//        poi2Html("xls/test.xlsx", "xls", "xls/img", "img");
    }

    public static void pdf2HtmlByPassword(String filePath, String htmlDir, String imgDir, String imgWebPath, String password) throws IOException {
        File file = POIUtils.checkFileExists(filePath);
        htmlDir = POIUtils.dealTargetDir(htmlDir);
        imgDir = POIUtils.dealTargetDir(imgDir);
        imgWebPath = POIUtils.dealTargetDir(imgWebPath);
//        PDF2Image.pdf2Image(file, htmlDir, imgDir, imgWebPath, password);
    }

    /**
     * @param filePath   待转换的原文件路径
     * @param htmlDir    生成的html文件路径
     * @param imgDir     原文件的图片存放目录
     * @param imgWebPath web应用访问图片的系统路径
     * @throws Exception
     */
    public static void poi2Html(String filePath, String htmlDir, String imgDir, String imgWebPath, HttpServletResponse res) throws Exception {
        File file = POIUtils.checkFileExists(filePath);
        htmlDir = POIUtils.dealTargetDir(htmlDir);
        imgDir = POIUtils.dealTargetDir(imgDir);
        imgWebPath = POIUtils.dealTargetDir(imgWebPath);
        byte[] data = File2byte.getBytes(file);
        if (filePath.endsWith("ppt") || filePath.endsWith("pptx")) {
            PPT2Image.ppt2Html(file, htmlDir, imgDir, imgWebPath, res);
        } else if (filePath.endsWith("doc") || filePath.endsWith("docx")) {
//            Word2Html.word2Html(file, htmlDir, imgDir, imgWebPath,res);

            if(filePath.endsWith("doc")){
                File2HtmlUtil.word2003ToHtml(data, res);
            }else if(filePath.endsWith("docx")){
                File2HtmlUtil.word2007ToHtml(data,res);
            }
        } else if (filePath.endsWith("xls") || filePath.endsWith("xlsx")) {
            Excel2Html.xls2Html(file, htmlDir, imgDir, imgWebPath,res);
        } else if (filePath.endsWith("pdf")) {
            PDF2Image.pdf2Image(file, htmlDir, imgDir, imgWebPath, null,res);
        } else if (filePath.endsWith("png") || filePath.endsWith("jpg") || filePath.endsWith("jpeg") || filePath.endsWith("gif")) {
            File2HtmlUtil.image2Html(data,res);
//        }else if(filePath.endsWith("mp4") || filePath.endsWith("ogv") || filePath.endsWith("webm")){
//            //视频页面预览和上面有些不一样，上面都是把文件转换成二进制或者流，而这个因为用了html5的<video/>标签，所以参数filePath得是能直接获取到视频的链接
//            String videoPath = "http://localhost:8989/video?file=" + file;
//            File2HtmlUtil.mp42Html(res,videoPath);
        } else {//如果文件类型上面没有，就直接下载
            res.setHeader("content-type", "multipart/form-data");
            res.setHeader("Content-Disposition", "attachment;filename=" +
                    new String( file.getName().getBytes("UTF-8"), "ISO8859-1" ));

            try (OutputStream os = res.getOutputStream()){
                if(LogicUtil.isNotEmpty(data)){
                    os.write(data);
                }
                os.flush();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }

    /**
     * @param file       待转换的原文件
     * @param htmlDir    生成的html文件路径
     * @param imgDir     原文件的图片存放目录
     * @param imgWebPath web应用访问图片的系统路径
     * @throws Exception
     */
//    public static void poi2Html(File file, String htmlDir, String imgDir, String imgWebPath) throws Exception {
//        String filePath = file.getPath();
//        if (!file.exists()) {
//            throw new RuntimeException("file not exists ![filepath = " + filePath + "]");
//        }
//
//        htmlDir = POIUtils.dealTargetDir(htmlDir);
//        imgDir = POIUtils.dealTargetDir(imgDir);
//        imgWebPath = POIUtils.dealTargetDir(imgWebPath);
//        poi2Html(filePath, htmlDir, imgDir, imgWebPath);
//    }
}
