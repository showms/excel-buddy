package com.funny.excelbuddy.export.dto;

/**
 * @author: 魏云飞 (weiyunfei@rd.keytop.com.cn)
 * @desc: .
 * @date: 2023/7/13.
 */
public class ExportCacheStatusDTO {
    private String state;
    private String progress;
    private String excelName;

    public String getState() {
        return state;
    }

    public void setState(String state) {
        this.state = state;
    }

    public String getProgress() {
        return progress;
    }

    public void setProgress(String progress) {
        this.progress = progress;
    }

    public String getExcelName() {
        return excelName;
    }

    public void setExcelName(String excelName) {
        this.excelName = excelName;
    }
}
