package com.funny.excelbuddy.export.config;

import com.alibaba.fastjson.JSON;
import com.github.pagehelper.PageInfo;
import com.google.common.collect.Maps;
import com.funny.excelbuddy.export.dto.PageDTO;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellRangeAddress;

import java.util.HashMap;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.function.BiConsumer;
import java.util.function.Function;

/**
 * @author: 魏云飞 (weiyunfei@rd.keytop.com.cn)
 * @desc: 导出配置
 * @date: 2023/7/17
 */
public class ExportConfig<T extends PageDTO, R> {

    /**
     * excel导出模板路径
     */
    private String exportTemplatePath;

    /**
     * 导出excel文件名称
     */
    private String exportName;

    /**
     * sheet导出配置
     * 适配常规模式下的单sheet导出配置
     */
    private SheetExportConfig<T, R> sheetExportConfig;

    /**
     * 多sheet导出配置
     * 常规模式下的单sheet导出配置会作为sheetExportConfigs的第一个元素，也就是sheetIndex=0
     */
    private Map<Integer, SheetExportConfig<? extends PageDTO, ?>> sheetExportConfigs;

    /**
     * 是否上传云存储，默认开启，以防分布式环境下下载失败
     */
    private boolean uploadCloudStore = true;

    /**
     * 上传云存储的临时文件会在缓存失效后自动删除
     */
    private boolean removeCloudStoreFileAfterExpire = true;

    /**
     * 默认异步执行
     */
    private boolean executeAsync = true;

    /**
     * 当前语言环境
     * 默认简体中文
     */
    private Locale locale = Locale.CHINA;

    public ExportConfig() {
        this.sheetExportConfig = new SheetExportConfig<>();
        this.sheetExportConfigs = Maps.newHashMap();
        sheetExportConfigs.put(0, sheetExportConfig);
    }

    public String getExportTemplatePath() {
        return exportTemplatePath;
    }

    public void setExportTemplatePath(String exportTemplatePath) {
        this.exportTemplatePath = exportTemplatePath;
    }

    public String getExportName() {
        return exportName;
    }

    public void setExportName(String exportName) {
        this.exportName = exportName;
    }

    public boolean isUploadCloudStore() {
        return uploadCloudStore;
    }

    public void setUploadCloudStore(boolean uploadCloudStore) {
        this.uploadCloudStore = uploadCloudStore;
    }

    public Locale getLocale() {
        return locale;
    }

    public void setLocale(Locale locale) {
        this.locale = locale;
    }

    public boolean isRemoveCloudStoreFileAfterExpire() {
        return removeCloudStoreFileAfterExpire;
    }

    public void setRemoveCloudStoreFileAfterExpire(boolean removeCloudStoreFileAfterExpire) {
        this.removeCloudStoreFileAfterExpire = removeCloudStoreFileAfterExpire;
    }

    public SheetExportConfig<T, R> getSheetExportConfig() {
        return sheetExportConfig;
    }

    public void setSheetExportConfig(SheetExportConfig<T, R> sheetExportConfig) {
        this.sheetExportConfig = sheetExportConfig;
    }

    public Function<T, PageInfo<R>> getGetExportDataSupplier() {
        return this.sheetExportConfig.getGetExportDataSupplier();
    }

    public void setGetExportDataSupplier(Function<T, PageInfo<R>> getExportDataSupplier) {
        this.sheetExportConfig.setGetExportDataSupplier(getExportDataSupplier);
    }

    public BiConsumer<Row, R> getFillExcelExportContentConsumer() {
        return this.sheetExportConfig.getFillExcelExportContentConsumer();
    }

    public void setFillExcelExportContentConsumer(BiConsumer<Row, R> fillExcelExportContentConsumer) {
        this.sheetExportConfig.setFillExcelExportContentConsumer(fillExcelExportContentConsumer);
    }

    public Integer getSourceDataRowIndex() {
        return this.sheetExportConfig.getSourceDataRowIndex();
    }

    public void setSourceDataRowIndex(Integer sourceDataRowIndex) {
        this.sheetExportConfig.setSourceDataRowIndex(sourceDataRowIndex);
    }

    public Integer getFillDataStartRowIndex() {
        return this.sheetExportConfig.getFillDataStartRowIndex();
    }

    public void setFillDataStartRowIndex(Integer fillDataStartRowIndex) {
        this.sheetExportConfig.setFillDataStartRowIndex(fillDataStartRowIndex);
    }

    public boolean isExportDataPageEnabled() {
        return this.sheetExportConfig.isExportDataPageEnabled();
    }

    public void setExportDataPageEnabled(boolean exportDataPageEnabled) {
        this.sheetExportConfig.setExportDataPageEnabled(exportDataPageEnabled);
    }

    public List<CellRangeAddress> getMergedRegionInfoList() {
        return this.sheetExportConfig.getMergedRegionInfoList();
    }

    public void setMergedRegionInfoList(List<CellRangeAddress> mergedRegionInfoList) {
        this.sheetExportConfig.setMergedRegionInfoList(mergedRegionInfoList);
    }

    public Integer getExportPageSize() {
        return this.sheetExportConfig.getExportPageSize();
    }

    public void setExportPageSize(Integer exportPageSize) {
        this.sheetExportConfig.setExportPageSize(exportPageSize);
    }

    @Override
    public String toString() {
        Map<String, Object> dataMap = new HashMap<String, Object>(16) {{
            put("exportTemplatePath", getExportTemplatePath());
            put("exportName", getExportName());
            put("uploadCloudStore", uploadCloudStore);
        }};
        return JSON.toJSONString(dataMap);
    }

    public Map<Integer, SheetExportConfig<? extends PageDTO, ?>> getSheetExportConfigs() {
        return sheetExportConfigs;
    }

    public void setSheetExportConfigs(Map<Integer, SheetExportConfig<? extends PageDTO, ?>> sheetExportConfigs) {
        this.sheetExportConfigs = sheetExportConfigs;
    }

    /**
     * 添加Sheet配置
     *
     * @param sheetIndex
     * @param functionalSheetExportConfig
     * @param <T>
     * @param <R>
     */
    public <T extends PageDTO, R> void addSheetConfig(int sheetIndex, FunctionalSheetExportConfig<T, R> functionalSheetExportConfig) {
        sheetExportConfigs.put(sheetIndex, functionalSheetExportConfig.apply());
    }

    /**
     * 获取Sheet配置
     *
     * @param sheetIndex Sheet索引
     * @return Sheet配置
     */
    @SuppressWarnings("unchecked")
    public <T extends PageDTO, R> SheetExportConfig<T, R> getSheetConfig(int sheetIndex) {
        return (SheetExportConfig<T, R>) sheetExportConfigs.get(sheetIndex);
    }

    public boolean isExecuteAsync() {
        return executeAsync;
    }

    public void setExecuteAsync(boolean executeAsync) {
        this.executeAsync = executeAsync;
    }
}
