package com.funny.excelbuddy.export.config;

import com.funny.excelbuddy.export.dto.PageDTO;

/**
 * @author: 魏云飞 (weiyunfei@rd.keytop.com.cn)
 * @desc:
 * @date: 2025/3/24
 */
@FunctionalInterface
public interface FunctionalExportConfig<T extends PageDTO, R> {

    /**
     * 获取导出配置
     *
     * @return
     */
    ExportConfig<T, R> apply();
}
