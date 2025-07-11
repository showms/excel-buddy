package com.funny.excelbuddy.export.utils;

import com.alibaba.fastjson.JSON;
import com.github.pagehelper.PageInfo;
import com.google.common.collect.ImmutableMap;
import com.google.common.collect.Maps;
import com.google.common.util.concurrent.ThreadFactoryBuilder;
import com.funny.excelbuddy.export.ApplicationContextHolder;
import com.funny.excelbuddy.export.config.ExportConfig;
import com.funny.excelbuddy.export.config.FunctionalExportConfig;
import com.funny.excelbuddy.export.config.SheetExportConfig;
import com.funny.excelbuddy.export.dto.PageDTO;
import com.funny.excelbuddy.export.service.RedisCacheService;
import com.sun.javaws.progress.Progress;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.slf4j.MDC;
import org.springframework.context.i18n.LocaleContextHolder;
import org.springframework.util.CollectionUtils;

import javax.servlet.http.HttpServletResponse;
import java.io.*;
import java.nio.charset.StandardCharsets;
import java.util.Map;
import java.util.Optional;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;
import java.util.concurrent.TimeUnit;

/**
 * @author: 魏云飞 (weiyunfei@rd.keytop.com.cn)
 * @desc: .
 * @date: 2018/3/15.
 */
public class ExportHelper {

    private static final Logger log = LoggerFactory.getLogger(ExportHelper.class);

    private static final String CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
    private static final String HEADER_NAME = "Content-disposition";
    private static final String HEADER_VALUE = "attachment;filename=";
    private static final String SUFFIX = ".xlsx";
    private static final String UPLOAD_CLOUD_STORE_DIRECTORY = "export";

    private static final String EXPORT_CONFIG_PREFIX = "export.config.";
    private static final String EXPORT_PROCESS_PREFIX = "export.process.";

    //导出结果缓存时间，单位分钟
    private static final Integer EXPORT_RESULT_CACHE_TIMEOUT = 60 * 2;

    private static Map<String, SXSSFWorkbookInfo> sXSSFWorkbookCacheMap = Maps.newHashMap();

    private static ExecutorService executorService = Executors.newFixedThreadPool(30, new ThreadFactoryBuilder().setNameFormat("export-pool-%d").build());

    /**
     * 导出任务
     * pageRequest为空，会从sheetConfig中获取
     *
     * @param taskId
     * @param functionalExportConfig
     */
    public static void exportTask(String taskId, FunctionalExportConfig functionalExportConfig) {
        ExportHelper.exportTask(taskId, null, functionalExportConfig);
    }

    /**
     * 导出任务
     *
     * @param taskId
     * @param pageRequest
     * @param functionalExportConfig
     * @param <T>
     * @param <R>
     */
    public static <T extends PageDTO, R> void exportTask(String taskId, T pageRequest, FunctionalExportConfig functionalExportConfig) {
        ExportConfig exportConfig = functionalExportConfig.apply();
        exportConfig.setLocale(LocaleContextHolder.getLocale());

        if (exportConfig.isExecuteAsync()) {
            Map<String, String> copyOfContextMap = MDC.getCopyOfContextMap();
            executorService.execute(() -> {
                try {
                    MDC.setContextMap(copyOfContextMap);
                    ExportHelper.executeExport(taskId, pageRequest, exportConfig);
                } catch (Exception e) {
                    log.error("导出失败，异常原因：{}", e.getMessage(), e);
                } finally {
                }
            });
            log.info("导出任务已提交，任务ID：{}", taskId);
        } else {
            ExportHelper.executeExport(taskId, pageRequest, exportConfig);
        }
    }

    /**
     * 执行导出
     *
     * @param taskId
     * @param pageRequest
     * @param exportConfig
     * @param <T>
     * @param <R>
     */
    private static <T extends PageDTO, R> void executeExport(String taskId, T pageRequest, ExportConfig exportConfig) {

        log.info("导出数据 开始，任务ID：{}", taskId);

        //初始化工作薄
        log.info("初始化工作薄{},{}", exportConfig.getExportTemplatePath(), exportConfig.getExportName());
        SXSSFWorkbook resultWB = ExcelUtil.initWorkbook(exportConfig.getExportTemplatePath());

        //调整为按sheet分别读取
        Map<Integer, SheetExportConfig<? extends PageDTO, ?>> sheetExportConfigs = exportConfig.getSheetExportConfigs();
        sheetExportConfigs.forEach((index, sheetExportConfig) -> {
            ExportHelper.fillData(taskId, resultWB, index, Optional.ofNullable(pageRequest).orElse((T) sheetExportConfig.getPageRequest()), exportConfig.getSheetConfig(index));
        });

        if (exportConfig.isUploadCloudStore()) {
            //上传云存储
            ExportHelper.uploadCloudStore(resultWB, taskId);
            ExportHelper.cacheExportConfig(resultWB, taskId, exportConfig);
        } else {
            ExportHelper.storeExportSXSSFWorkbook(taskId, exportConfig.getExportName(), resultWB);
        }

        ExportHelper.updateExportProgress(taskId, Boolean.TRUE, "已全部导出");

        log.info("导出数据 结束，任务ID：{}", taskId);
    }

    /**
     * 填充数据
     *
     * @param taskId
     * @param resultWB
     * @param index
     * @param pageRequest
     * @param sheetExportConfig
     * @param <T>
     * @param <R>
     */
    private static <T extends PageDTO, R> void fillData(String taskId, SXSSFWorkbook resultWB, int index, T pageRequest, SheetExportConfig<T, R> sheetExportConfig) {
        //数据复制格式的行索引
        int fillDataStartRowIndex = sheetExportConfig.getFillDataStartRowIndex();
        pageRequest.setPageSize(sheetExportConfig.getExportPageSize());

        //数据模板行
        Row srcDataRow = ExcelUtil.getSrcRow(resultWB, index, sheetExportConfig.getSourceDataRowIndex());

        for (int pageNum = 1; ; pageNum++) {
            pageRequest.setPageNo(pageNum);
            PageInfo<R> exportDataPageInfo = sheetExportConfig.getGetExportDataSupplier().apply(pageRequest);
            long total = exportDataPageInfo.getSize();
            if (pageNum == 1 && total == 0) {
                ExportHelper.updateExportProgress(taskId, Boolean.FALSE, "没有记录");
                return;
            } else if (total == 0) {
                break;
            }
            log.info("sheetIndex：{}，总记录数：{}，页数：{}，第{}页数据读取完毕", index, exportDataPageInfo.getTotal(), exportDataPageInfo.getPages(), pageNum);

            for (R data : exportDataPageInfo.getList()) {
                Row destRow = ExcelUtil.getDestRow(resultWB, srcDataRow, index, fillDataStartRowIndex++);
                sheetExportConfig.getFillExcelExportContentConsumer().accept(destRow, data);
            }

            if (!CollectionUtils.isEmpty(sheetExportConfig.getMergedRegionInfoList())) {
                //添加合并单元格区
                sheetExportConfig.getMergedRegionInfoList().forEach(region -> {
                    resultWB.getSheetAt(index).addMergedRegion(region);
                });
            }

            if (!sheetExportConfig.isExportDataPageEnabled()) {
                //支持不翻页
                break;
            }
        }
    }

    /**
     * 缓存导出配置
     *
     * @param resultWB
     * @param taskId
     * @param exportConfig
     */
    public static void cacheExportConfig(SXSSFWorkbook resultWB, String taskId, ExportConfig exportConfig) {
        log.info("有开启上传到云存储，开始缓存导出配置，有效期30分钟：{} {}", taskId, exportConfig.toString());
        //缓存临时文件 EXPORT_RESULT_CACHE_TIMEOUT
        ApplicationContextHolder.getBean(RedisCacheService.class).setEx(EXPORT_CONFIG_PREFIX.concat(taskId), exportConfig.toString(), EXPORT_RESULT_CACHE_TIMEOUT, TimeUnit.MINUTES);
        if (exportConfig.isRemoveCloudStoreFileAfterExpire()) {
            ExportHelper.removeCloudFileDelay(taskId, EXPORT_RESULT_CACHE_TIMEOUT);
        }
    }

    /**
     * 上传云存储
     *
     * @param resultWB
     * @param taskId
     */
    public static void uploadCloudStore(SXSSFWorkbook resultWB, String taskId) {
        String filePath = taskId + SUFFIX;
        File tempFile = null;
        try (FileOutputStream fos = new FileOutputStream(filePath)) {
            resultWB.write(fos);
            tempFile = new File(filePath);
            try (FileInputStream in = new FileInputStream(tempFile);) {
                CloudStoreUtil.uploadFile(UPLOAD_CLOUD_STORE_DIRECTORY, filePath, in);
            } catch (Exception e) {
                log.error("上传云存储失败，异常原因：{}", e.getMessage(), e);
            }
        } catch (Exception e) {
            log.error("resultWB写入到输出流失败，异常原因：{}", e.getMessage(), e);
        } finally {
            if (tempFile != null && tempFile.exists()) {
                tempFile.delete();
            }
            resultWB.dispose();
        }
    }

    /**
     * 下载EXCEL
     *
     * @param taskId
     * @param response
     */
    public static void downloadExcel(String taskId, HttpServletResponse response) {
        log.info("下载excel 开始：{}", taskId);
        OutputStream out = null;
        SXSSFWorkbook resultWB = null;
        try {
            //初始化工作薄
            SXSSFWorkbookInfo workbookInfo = sXSSFWorkbookCacheMap.get(taskId);
            if (workbookInfo == null) {
                log.error("没有excel工作簿，不处理");
                return;
            }
            resultWB = workbookInfo.workbook;
            //保存报表
            final String name = workbookInfo.getExcelName();
            out = response.getOutputStream();
            response.setContentType(CONTENT_TYPE);
            response.setHeader(HEADER_NAME, HEADER_VALUE + java.net.URLEncoder.encode(name, StandardCharsets.UTF_8.name()) + SUFFIX);
            resultWB.write(out);
            out.flush();
        } catch (IOException e) {
            log.error("下载excel失败，异常原因：{}", e.getMessage(), e);
        } finally {
            // dispose of temporary files backing this workbook on disk
            if (resultWB != null) {
                resultWB.dispose();
            }
            try {
                if (out != null) {
                    out.close();
                }
            } catch (IOException e) {
                log.error("下载excel失败，异常原因：{}", e.getMessage(), e);
            }
            log.info("下载excel 结束：{}", taskId);
        }
    }

    /**
     * 从云存储下载
     * 需要具体对接云服务sdk，比如oss cos azure
     *
     * @param taskId
     * @param response
     */
    public static void downloadExcelCloud(String taskId, HttpServletResponse response) {
        Optional<ExportConfig> optional = ApplicationContextHolder.getBean(RedisCacheService.class).get(EXPORT_CONFIG_PREFIX.concat(taskId), ExportConfig.class);
        if (!optional.isPresent()) {
            log.info("下载失败，文件已失效：{}", taskId);
            return;
        }
        CloudStoreUtil.downloadFile(UPLOAD_CLOUD_STORE_DIRECTORY, taskId + SUFFIX, optional.get().getExportName() + SUFFIX, response);
    }

    /**
     * 下载导出模板
     *
     * @param templatePath
     * @param exportName
     * @param response
     */
    public static void downloadExportTemplate(String templatePath, String exportName, HttpServletResponse response) {
        log.info("下载excel模板 开始：{}", templatePath);
        OutputStream out = null;
        SXSSFWorkbook resultWB = null;
        try {
            //初始化工作薄
            resultWB = ExcelUtil.initWorkbook(templatePath);

            //保存报表
            final String name = exportName;
            out = response.getOutputStream();
            response.setContentType(CONTENT_TYPE);
            response.setHeader(HEADER_NAME, HEADER_VALUE + java.net.URLEncoder.encode(name, StandardCharsets.UTF_8.name()) + SUFFIX);
            resultWB.write(out);
            out.flush();
        } catch (IOException e) {
            log.error("下载excel模板失败：{}", e.getMessage(), e);
        } finally {
            // dispose of temporary files backing this workbook on disk
            if (resultWB != null) {
                resultWB.dispose();
            }
            try {
                if (out != null) {
                    out.close();
                }
            } catch (IOException e) {
                log.error("关闭输出流失败：{}", e.getMessage(), e);
            }
            log.info("下载excel 结束：{}", templatePath);
        }
    }

    /**
     * 获取导出进度
     * 分布式服务需要从缓存读取
     * 若是单机服务，从内存中读取
     *
     * @param taskId
     * @return
     */
    public static Progress getExportProgress(String taskId) {
        Optional<Progress> optional = ApplicationContextHolder.getBean(RedisCacheService.class).get(EXPORT_PROCESS_PREFIX.concat(taskId), Progress.class);
        return optional.orElse(null);
    }

    /**
     * 更新导出进度
     * 分布式服务需要将导出进度进行统一缓存
     * 若是单机服务，那可以直接保存在内存中
     *
     * @param taskId
     * @param success
     * @param msg
     */
    public static void updateExportProgress(String taskId, Boolean success, String msg) {
        Progress exportProgress = new Progress();
        exportProgress.setResult(success);
        exportProgress.setMsg(msg);
        ApplicationContextHolder.getBean(RedisCacheService.class).setEx(EXPORT_PROCESS_PREFIX.concat(taskId), JSON.toJSONString(exportProgress), EXPORT_RESULT_CACHE_TIMEOUT, TimeUnit.MINUTES);
    }

    /**
     * 延迟删除云储存临时文件
     * 有条件的可以借助MQ延迟消息，删除云存储中的临时文件
     * 如果云存储支持设置某个目录过期策略也可以
     * 当然也可以简单搞个定时任务删
     *
     * @param taskId
     * @param delayMinutes
     */
    public static void removeCloudFileDelay(String taskId, Integer delayMinutes) {
        try {
            //ProducerSendMessageFacade.sendPulsarDelayMsg(DataGuard.MAJOR, TaskConstants.REMOVE_CLOUD_FILE_TOPIC,
            //      JSON.toJSONString(ImmutableMap.of("fileDirectory", UPLOAD_CLOUD_STORE_DIRECTORY, "fileName", taskId + SUFFIX)), 60 * (delayMinutes + 1));
        } catch (Exception e) {
            log.error("发送删除云储存临时文件延迟消息出现异常", e);
        }
    }

    /**
     * 清除内存
     *
     * @param taskId
     */
    public static void cleanExportCacheByKey(String taskId) {
        sXSSFWorkbookCacheMap.remove(taskId);
    }

    /**
     * 存储导出工作薄
     *
     * @param taskId
     * @param resultWB
     */
    public static void storeExportSXSSFWorkbook(String taskId, String excelName, SXSSFWorkbook resultWB) {
        sXSSFWorkbookCacheMap.put(taskId, new SXSSFWorkbookInfo(excelName, resultWB));
    }

    /**
     * excel工作簿信息
     */
    public static class SXSSFWorkbookInfo {
        private String excelName;
        private SXSSFWorkbook workbook;

        public SXSSFWorkbookInfo() {
        }

        public SXSSFWorkbookInfo(String excelName, SXSSFWorkbook workbook) {
            this.excelName = excelName;
            this.workbook = workbook;
        }

        public SXSSFWorkbook getWorkbook() {
            return workbook;
        }

        public void setWorkbook(SXSSFWorkbook workbook) {
            this.workbook = workbook;
        }

        public String getExcelName() {
            return excelName;
        }

        public void setExcelName(String excelName) {
            this.excelName = excelName;
        }
    }

    /**
     * 进度类
     */
    public static class Progress {
        private Boolean result;
        private String msg;

        public Boolean getResult() {
            return result;
        }

        public void setResult(Boolean result) {
            this.result = result;
        }

        public String getMsg() {
            return msg;
        }

        public void setMsg(String msg) {
            this.msg = msg;
        }
    }
}
