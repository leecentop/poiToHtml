package top.lllyl2012.html.service;

import org.springframework.http.ResponseEntity;

import javax.servlet.http.HttpServletResponse;

/**
 * @Author: volume
 * @Description:
 * @CreateDate: 2019/6/15 11:26
 */
public interface ResourceService {
    void toHtml(HttpServletResponse response,String fileName);

    ResponseEntity<?> showPhotoDoc(HttpServletResponse response, String photo);

    ResponseEntity<?> showPhotoDoc(HttpServletResponse response, String dir, String photo);

    ResponseEntity<?> showPhotoDocx(HttpServletResponse response, String dir, String photo);
}
