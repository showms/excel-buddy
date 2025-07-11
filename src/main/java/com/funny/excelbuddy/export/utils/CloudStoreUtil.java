package com.funny.excelbuddy.export.utils;

import javax.servlet.http.HttpServletResponse;
import java.io.FileInputStream;

/**
 * @author: 魏云飞 (weiyunfei@rd.keytop.com.cn)
 * @desc: 云储存工具类，对接云存储服务，比如oss cos azure
 * @date: 2025/7/11
 */
public class CloudStoreUtil {

    /**
     * 上传
     *
     * @param fileDirectory  目录
     * @param uploadFileName 上传文件名
     * @param in             文件流
     */
    public static void uploadFile(String fileDirectory, String uploadFileName, FileInputStream in) {

    }

    /**
     * 下载
     *
     * @param fileDirectory 文件目录
     * @param cloudFileName 云存储中的文件名称
     * @param fileName      下载的文件名称
     * @param response
     */
    public static void downloadFile(String fileDirectory, String cloudFileName, String fileName, HttpServletResponse response) {

    }
}
