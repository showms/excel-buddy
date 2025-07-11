package com.funny.excelbuddy.export.config;

import com.github.pagehelper.PageInfo;
import com.google.common.collect.Lists;
import com.funny.excelbuddy.export.dto.PageDTO;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellRangeAddress;

import java.util.List;
import java.util.function.BiConsumer;
import java.util.function.Function;

/**
 * @author: 魏云飞 (weiyunfei@rd.keytop.com.cn)
 * @desc: 单个sheet导出配置
 * @date: 2025/3/21
 */
public class SheetExportConfig<T extends PageDTO, R> {

    private T pageRequest;

    /**
     * 数据复制格式的行索引，从0开始
     * 默认等于1，表示复制rowIndex=1的单元格格式
     */
    private Integer sourceDataRowIndex = 1;

    /**
     * 数据开始填充数据的行索引，从0开始
     * 默认等于1，表示从rowIndex=1的地方开始填充数据
     */
    private Integer fillDataStartRowIndex = 1;

    /**
     * 获取导出数据
     */
    private Function<T, PageInfo<R>> getExportDataSupplier;

    /**
     * 导出数据开启翻页，默认开启
     */
    private boolean exportDataPageEnabled = true;

    /**
     * 填充excel导出内容
     */
    private BiConsumer<Row, R> fillExcelExportContentConsumer;

    /**
     * 合并单元格信息
     */
    private List<CellRangeAddress> mergedRegionInfoList = Lists.newArrayList();

    /**
     * 分頁大小
     */
    private Integer exportPageSize = 1024;

    public Integer getSourceDataRowIndex() {
        return sourceDataRowIndex;
    }

    public void setSourceDataRowIndex(Integer sourceDataRowIndex) {
        this.sourceDataRowIndex = sourceDataRowIndex;
    }

    public Integer getFillDataStartRowIndex() {
        return fillDataStartRowIndex;
    }

    public void setFillDataStartRowIndex(Integer fillDataStartRowIndex) {
        this.fillDataStartRowIndex = fillDataStartRowIndex;
    }

    public Function<T, PageInfo<R>> getGetExportDataSupplier() {
        return getExportDataSupplier;
    }

    public void setGetExportDataSupplier(Function<T, PageInfo<R>> getExportDataSupplier) {
        this.getExportDataSupplier = getExportDataSupplier;
    }

    public boolean isExportDataPageEnabled() {
        return exportDataPageEnabled;
    }

    public void setExportDataPageEnabled(boolean exportDataPageEnabled) {
        this.exportDataPageEnabled = exportDataPageEnabled;
    }

    public BiConsumer<Row, R> getFillExcelExportContentConsumer() {
        return fillExcelExportContentConsumer;
    }

    public void setFillExcelExportContentConsumer(BiConsumer<Row, R> fillExcelExportContentConsumer) {
        this.fillExcelExportContentConsumer = fillExcelExportContentConsumer;
    }

    public List<CellRangeAddress> getMergedRegionInfoList() {
        return mergedRegionInfoList;
    }

    public void setMergedRegionInfoList(List<CellRangeAddress> mergedRegionInfoList) {
        this.mergedRegionInfoList = mergedRegionInfoList;
    }

    public Integer getExportPageSize() {
        return exportPageSize;
    }

    public void setExportPageSize(Integer exportPageSize) {
        this.exportPageSize = exportPageSize;
    }

    public T getPageRequest() {
        return pageRequest;
    }

    public void setPageRequest(T pageRequest) {
        this.pageRequest = pageRequest;
    }
}
