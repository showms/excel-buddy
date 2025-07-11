package com.funny.excelbuddy.export.controller;

import com.funny.excelbuddy.export.res.ApiResponse;
import com.funny.excelbuddy.export.utils.ExportHelper;
import org.springframework.web.bind.annotation.*;

import javax.servlet.http.HttpServletResponse;

/**
 * @author: 魏云飞 (weiyunfei@rd.keytop.com.cn)
 * @desc:
 * @date: 2023/7/29
 */
@RestController
@RequestMapping("/service/excel")
public class ExportProcessController {

    /**
     * 查询导出状态
     *
     * @param taskId
     * @return
     */
    @ResponseBody
    @PostMapping(value = "/getExportProgress")
    public ApiResponse<ExportHelper.Progress> queryExportState(String taskId) {
        ExportHelper.Progress exportProgress = ExportHelper.getExportProgress(taskId);
        return ApiResponse.success("success", exportProgress);
    }

    /**
     * 下载EXCEL
     *
     * @param taskId
     * @param response
     */
    @GetMapping(value = "/downloadExcel")
    public void downloadExcel(String taskId, HttpServletResponse response) {
        ExportHelper.downloadExcel(taskId, response);
    }

    /**
     * 下载EXCEL
     * 从云存储下载
     *
     * @param taskId
     * @param response
     */
    @GetMapping(value = "/downloadExcelCloud")
    public void downloadExcelCloud(String taskId, HttpServletResponse response) {
        ExportHelper.downloadExcelCloud(taskId, response);
    }
}
