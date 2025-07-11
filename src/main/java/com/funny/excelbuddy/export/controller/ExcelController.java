package com.funny.excelbuddy.export.controller;

import com.alibaba.fastjson.JSONArray;
import com.funny.excelbuddy.export.config.ExportConfig;
import com.funny.excelbuddy.export.dto.PageDTO;
import com.funny.excelbuddy.export.res.ApiResponse;
import com.funny.excelbuddy.export.utils.ExportHelper;
import com.funny.excelbuddy.export.utils.ImportHelper;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestPart;
import org.springframework.web.bind.annotation.ResponseBody;
import org.springframework.web.multipart.MultipartFile;

import java.util.UUID;

/**
 * @author: 魏云飞 (weiyunfei@rd.keytop.com.cn)
 * @desc:
 * @date: 2025/2/26
 */
public class ExcelController<T extends PageDTO> {

    private static final Logger log = LoggerFactory.getLogger(ExcelController.class);

    /**
     * 导出
     *
     * @param pageRequest
     * @return
     */
    @ResponseBody
    @PostMapping(value = "/export")
    public ApiResponse export(@RequestBody T pageRequest) {
        String taskId = UUID.randomUUID().toString();
        ExportHelper.exportTask(taskId, pageRequest, () -> this.getExportConfig());
        return ApiResponse.success("操作成功", taskId);
    }

    /**
     * 导入
     *
     * @param file
     * @return
     * @throws Exception
     */
    @PostMapping("/import")
    public void importUser(@RequestPart("file") MultipartFile file) throws Exception {
        this.getImportData(ImportHelper.readMultipartFile(file));
    }

    /**
     * 导出配置，需要导出功能时需重写该方法
     * @param <T> 请求实体
     * @param <R> 结果实体
     * @return
     */
    protected <T extends PageDTO, R> ExportConfig<T, R> getExportConfig() {
        return null;
    }

    /**
     * 获取导入数据，需要导入功能时需要重写该方法
     *
     * @param data
     */
    protected void getImportData(JSONArray data) {
        log.info("导入数据：{}", JSONArray.toJSONString(data));
    }
}