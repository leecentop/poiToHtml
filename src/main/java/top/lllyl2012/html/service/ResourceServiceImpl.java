package top.lllyl2012.html.service;

import com.tfd.poi2html.POI2Html;
import com.tfd.poi2html.PPT2Image;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.core.io.Resource;
import org.springframework.core.io.ResourceLoader;
import org.springframework.http.HttpEntity;
import org.springframework.http.HttpHeaders;
import org.springframework.http.HttpMethod;
import org.springframework.http.ResponseEntity;
import org.springframework.stereotype.Service;
import org.springframework.util.ResourceUtils;
import top.lllyl2012.html.utils.File2HtmlUtil;
import top.lllyl2012.html.utils.File2byte;
import top.lllyl2012.html.utils.LogicUtil;

import javax.servlet.http.HttpServletResponse;
import java.io.*;
import java.nio.file.Paths;

/**
 * @Author: volume
 * @CreateDate: 2019/6/15 11:27
 */
@Service
public class ResourceServiceImpl implements ResourceService {

    private final String DOC = ".doc";
    private final String DOCX = ".docx";
    private final String PDF = ".pdf";
    private final String PNG = ".png";
    private final String JPG = ".jpg";
    private final String JPEG = ".jpeg";
    private final String GIF = ".gif";
    private final String XLS = ".xls";
    private final String XLSX = ".xlsx";
    private final String MP4 = ".mp4";
    private final String OGV = ".ogv";
    private final String WEBM = ".webm";

    private final String PPT = ".ppt";
    private final String PPTX = ".pptx";


    @Value("${resource-path}")
    private String resourcePath;


    private final ResourceLoader resourceLoader;

    @Autowired
    public ResourceServiceImpl(ResourceLoader resourceLoader) {
        this.resourceLoader = resourceLoader;
    }

    @Override
    public void toHtml(HttpServletResponse res,String fileName) {
        res.setContentType("text/html");

        FileInputStream fis = null;
        try {
            //获得本地的文件流
//            String fileName = "b.doc";
//            File file = ResourceUtils.getFile("classpath:static/"+fileName);
            File file = ResourceUtils.getFile(resourcePath+fileName);
//            byte[] data = File2byte.getBytes(file);
//
//            int suffixIndex = fileName.lastIndexOf(".");
//
//            if(DOC.equals(fileName.substring(suffixIndex))){
//                File2HtmlUtil.word2003ToHtml(data,res);
//            }else if(DOCX.equals(fileName.substring(suffixIndex))){
//                File2HtmlUtil.word2007ToHtml(data,res);
//            }else  if(PDF.equals(fileName.substring(suffixIndex))){
//                File2HtmlUtil.pdf2Html(data,res);
//            }else if(PNG.equals(fileName.substring(suffixIndex)) || JPG.equals(fileName.substring(suffixIndex))
//                    || JPEG.equals(fileName.substring(suffixIndex)) || GIF.equals(fileName.substring(suffixIndex))){
//                File2HtmlUtil.image2Html(data,res);
//            }else if(XLS.equals(fileName.substring(suffixIndex))){
//                InputStream in = new ByteArrayInputStream(data);
//                File2HtmlUtil.PoiExcelToHtml(in,res);
//            }else if(XLSX.equals(fileName.substring(suffixIndex))){
//                InputStream in = new ByteArrayInputStream(data);
//                File2HtmlUtil.ExcelToHtml(in,res);
//            }else if(MP4.equals(fileName.substring(suffixIndex)) || OGV.equals(fileName.substring(suffixIndex)) || WEBM.equals(fileName.substring(suffixIndex))){
//                //视频页面预览和上面有些不一样，上面都是把文件转换成二进制或者流，而这个因为用了html5的<video/>标签，所以参数filePath得是能直接获取到视频的链接
//                String filePath = "http://localhost:8989/video?file=" + fileName;
//                File2HtmlUtil.mp42Html(res,filePath);
//            }else if(PPT.equals(fileName.substring(suffixIndex))){
//                File2HtmlUtil.getTextFromPPT(file.getPath());
//                poi2Html("xls/test.xls", "xls", "xls/img", "img");
                POI2Html.poi2Html(file.getPath(), "ppt", File2HtmlUtil.getRoot() + "/image/word/media/", "",res);
//                PPT2Image.ppt2Html(file, "ppt", File2HtmlUtil.getRoot() + "/image/word/media/", "",res);
//            }else  if(PPTX.equals(fileName.substring(suffixIndex))){
////                File2HtmlUtil.getTextFromPPTX(file.getPath());
//            }else{//如果文件类型上面没有，就直接下载
//                res.setHeader("content-type", "multipart/form-data");
//                res.setHeader("Content-Disposition", "attachment;filename=" +
//                        new String( fileName.getBytes("UTF-8"), "ISO8859-1" ));
//
//                try (OutputStream os = res.getOutputStream()){
//                    if(LogicUtil.isNotEmpty(data)){
//                        os.write(data);
//                    }
//                    os.flush();
//                } catch (IOException e) {
//                    e.printStackTrace();
//                }
//            }
        } catch (Exception e) {
            e.printStackTrace();
        } catch (Throwable throwable) {
            throwable.printStackTrace();
        }
    }

    @Override
    public ResponseEntity<?> showPhotoDoc(HttpServletResponse response, String photo) {
        String root = File2HtmlUtil.getRoot()+"/image/word/media";
        Resource resource = resourceLoader.getResource("file:" + Paths.get(root, photo));
        return ResponseEntity.ok(resource);
    }

    @Override
    public ResponseEntity<?> showPhotoDoc(HttpServletResponse response,String dir, String photo) {
        String root = File2HtmlUtil.getRoot()+"/image/word/media"+File.separator + dir;
        Resource resource = resourceLoader.getResource("file:" + Paths.get(root, photo));
        return ResponseEntity.ok(resource);
    }

    @Override
    public ResponseEntity<?> showPhotoDocx(HttpServletResponse response,String dir, String photo) {
        String root = File2HtmlUtil.getRoot()+"/image"+File.separator + dir +"/word/media";
        Resource resource = resourceLoader.getResource("file:" + Paths.get(root, photo));
        return ResponseEntity.ok(resource);
    }

}
