package com.funny.excelbuddy.export.controller;

import com.github.pagehelper.PageInfo;
import com.google.common.collect.Lists;
import com.google.common.collect.Maps;
import com.funny.excelbuddy.export.config.ExportConfig;
import com.funny.excelbuddy.export.config.SheetExportConfig;
import com.funny.excelbuddy.export.dto.BusinessExportDataDTO;
import com.funny.excelbuddy.export.dto.MergedRow;
import com.funny.excelbuddy.export.req.GetDataPageRequest;
import com.funny.excelbuddy.export.res.ApiResponse;
import com.funny.excelbuddy.export.service.BusinessService;
import com.funny.excelbuddy.export.utils.ExportHelper;
import org.apache.poi.ss.util.CellRangeAddress;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import java.util.List;
import java.util.Map;
import java.util.UUID;

/**
 * @author: 魏云飞 (weiyunfei@rd.keytop.com.cn)
 * @desc:
 * @date: 2025/2/27
 */
@RequestMapping("/account")
@RestController
public class BusinessController extends ExcelController<GetDataPageRequest> {

    @Autowired
    private BusinessService businessService;

    /**
     * 列表数据
     *
     * @param pageRequest
     * @return
     */
    @PostMapping("/listPage")
    public ApiResponse<PageInfo<BusinessExportDataDTO>> listPage(@RequestBody GetDataPageRequest pageRequest) {
        return ApiResponse.success("获取成功", this.businessService.listData(pageRequest));
    }

    /**
     * 导出方式1，继承ExcelController并重写getExportConfig方法
     * 导出配置，需要导出功能时需重写该方法
     *
     * @return
     */
    @Override
    protected ExportConfig<GetDataPageRequest, BusinessExportDataDTO> getExportConfig() {
        ExportConfig<GetDataPageRequest, BusinessExportDataDTO> exportConfig = new ExportConfig<>();
        exportConfig.setExportName("导出方式1");
        exportConfig.setExportTemplatePath("templates/business_export_template.xlsx");
        //导出数据查询过程中不翻页
        exportConfig.setExportDataPageEnabled(false);
        exportConfig.setUploadCloudStore(false);
        exportConfig.setGetExportDataSupplier(pageRequest -> this.businessService.listExportData(pageRequest));
        exportConfig.setFillExcelExportContentConsumer((destRow, data) -> {
            destRow.getCell(0).setCellValue(data.getUid());
            destRow.getCell(1).setCellValue(data.getName());
            destRow.getCell(2).setCellValue(data.getCardNo());
            destRow.getCell(3).setCellValue(data.getBalance());
        });
        return exportConfig;
    }

    /**
     * 导出方式2，无需继承ExcelController
     *
     * @param pageRequest
     * @return
     */
    @PostMapping("/export2")
    public ApiResponse export2(@RequestBody GetDataPageRequest pageRequest) {
        String taskId = UUID.randomUUID().toString();
        ExportHelper.exportTask(taskId, pageRequest,
                () -> {
                    ExportConfig<GetDataPageRequest, BusinessExportDataDTO> exportConfig = new ExportConfig<>();
                    exportConfig.setExportName("导出方式2");
                    exportConfig.setExportTemplatePath("templates/business_export_template.xlsx");
                    //导出数据查询过程中不翻页
                    exportConfig.setExportDataPageEnabled(false);
                    exportConfig.setUploadCloudStore(false);
                    exportConfig.setGetExportDataSupplier(request -> this.businessService.listExportData(request));
                    exportConfig.setFillExcelExportContentConsumer((destRow, data) -> {
                        destRow.getCell(0).setCellValue(data.getUid());
                        destRow.getCell(1).setCellValue(data.getName());
                        destRow.getCell(2).setCellValue(data.getCardNo());
                        destRow.getCell(3).setCellValue(data.getBalance());
                    });
                    return exportConfig;
                });
        return ApiResponse.success("请求成功", taskId);
    }

    @PostMapping("/export3")
    public ApiResponse export3(@RequestBody GetDataPageRequest pageRequest) {
        String taskId = UUID.randomUUID().toString();
        ExportHelper.exportTask(taskId, pageRequest,
                () -> {
                    ExportConfig<GetDataPageRequest, BusinessExportDataDTO> exportConfig = new ExportConfig<>();
                    exportConfig.setExportName("导出方式3-含合并单元格");
                    exportConfig.setExportTemplatePath("templates/business_export_template.xlsx");
                    //导出数据查询过程中不翻页
                    exportConfig.setExportDataPageEnabled(false);
                    exportConfig.setUploadCloudStore(false);

                    exportConfig.setGetExportDataSupplier(request -> {
                        PageInfo<BusinessExportDataDTO> pageInfo = this.businessService.listExportData(request);

                        //提取要合并单元格的行索引
                        Map<String, MergedRow> rowIndexMap = Maps.newHashMap();
                        List<BusinessExportDataDTO> exportList = pageInfo.getList();
                        for (int i = 0; i < exportList.size(); i++) {
                            BusinessExportDataDTO dataDTO = exportList.get(i);
                            String key = dataDTO.getUid() + dataDTO.getName();
                            MergedRow mergedRow = rowIndexMap.get(key);
                            if (mergedRow == null) {
                                //因为有标题的缘故，合并单元格rowIndex是i+1
                                mergedRow = new MergedRow(i + 1, i + 1);
                                rowIndexMap.put(key, mergedRow);
                            } else {
                                //因为有标题的缘故，合并单元格rowIndex是i+1
                                mergedRow.setEndIndex(i + 1);
                            }
                        }

                        //合并单元格信息
                        List<CellRangeAddress> mergedRegionInfoList = Lists.newArrayList();
                        rowIndexMap.values().stream()
                                .filter(mergedRow -> mergedRow.getEndIndex() > mergedRow.getStartIndex())
                                .forEach(mergedRow -> {
                                    mergedRegionInfoList.add(new CellRangeAddress(mergedRow.getStartIndex(), mergedRow.getEndIndex(), 0, 0));
                                    mergedRegionInfoList.add(new CellRangeAddress(mergedRow.getStartIndex(), mergedRow.getEndIndex(), 1, 1));
                                });
                        exportConfig.setMergedRegionInfoList(mergedRegionInfoList);

                        return pageInfo;
                    });
                    exportConfig.setFillExcelExportContentConsumer((destRow, data) -> {
                        destRow.getCell(0).setCellValue(data.getUid());
                        destRow.getCell(1).setCellValue(data.getName());
                        destRow.getCell(2).setCellValue(data.getCardNo());
                        destRow.getCell(3).setCellValue(data.getBalance());
                    });
                    return exportConfig;
                });
        return ApiResponse.success("请求成功", taskId);
    }

    @PostMapping("/export4")
    public ApiResponse export4(@RequestBody GetDataPageRequest pageRequest) {
        String taskId = UUID.randomUUID().toString();
        ExportHelper.exportTask(taskId, pageRequest,
                () -> {
                    ExportConfig exportConfig = new ExportConfig();
                    exportConfig.setExportName("导出方式4-多sheet");
                    exportConfig.setExportTemplatePath("templates/business_export_multi_template.xlsx");
                    exportConfig.setUploadCloudStore(false);

                    //第一个sheet
                    exportConfig.addSheetConfig(0, () -> {
                        SheetExportConfig<GetDataPageRequest, BusinessExportDataDTO> sheetExportConfig = new SheetExportConfig<>();

                        sheetExportConfig.setPageRequest(pageRequest);
                        //导出数据查询过程中不翻页
                        sheetExportConfig.setExportDataPageEnabled(Boolean.FALSE);
                        sheetExportConfig.setGetExportDataSupplier(request -> {
                            PageInfo<BusinessExportDataDTO> pageInfo = this.businessService.listExportData(request);

                            //提取要合并单元格的行索引
                            Map<String, MergedRow> rowIndexMap = Maps.newHashMap();
                            List<BusinessExportDataDTO> exportList = pageInfo.getList();
                            for (int i = 0; i < exportList.size(); i++) {
                                BusinessExportDataDTO dataDTO = exportList.get(i);
                                String key = dataDTO.getUid() + dataDTO.getName();
                                MergedRow mergedRow = rowIndexMap.get(key);
                                if (mergedRow == null) {
                                    //因为有标题的缘故，合并单元格rowIndex是i+1
                                    mergedRow = new MergedRow(i + 1, i + 1);
                                    rowIndexMap.put(key, mergedRow);
                                } else {
                                    //因为有标题的缘故，合并单元格rowIndex是i+1
                                    mergedRow.setEndIndex(i + 1);
                                }
                            }

                            //合并单元格信息
                            List<CellRangeAddress> mergedRegionInfoList = Lists.newArrayList();
                            rowIndexMap.values().stream()
                                    .filter(mergedRow -> mergedRow.getEndIndex() > mergedRow.getStartIndex())
                                    .forEach(mergedRow -> {
                                        mergedRegionInfoList.add(new CellRangeAddress(mergedRow.getStartIndex(), mergedRow.getEndIndex(), 0, 0));
                                        mergedRegionInfoList.add(new CellRangeAddress(mergedRow.getStartIndex(), mergedRow.getEndIndex(), 1, 1));
                                    });
                            sheetExportConfig.setMergedRegionInfoList(mergedRegionInfoList);

                            return pageInfo;
                        });
                        sheetExportConfig.setFillExcelExportContentConsumer((destRow, data) -> {
                            destRow.getCell(0).setCellValue(data.getUid());
                            destRow.getCell(1).setCellValue(data.getName());
                            destRow.getCell(2).setCellValue(data.getCardNo());
                            destRow.getCell(3).setCellValue(data.getBalance());
                        });

                        return sheetExportConfig;
                    });

                    //第二个sheet
                    exportConfig.addSheetConfig(1, () -> {
                        SheetExportConfig<GetDataPageRequest, BusinessExportDataDTO> sheetExportConfig = new SheetExportConfig<>();
                        sheetExportConfig.setPageRequest(pageRequest);
                        //导出数据查询过程中不翻页
                        sheetExportConfig.setExportDataPageEnabled(Boolean.FALSE);
                        sheetExportConfig.setGetExportDataSupplier(request -> this.businessService.listExportData(request));
                        sheetExportConfig.setFillExcelExportContentConsumer((destRow, data) -> {
                            destRow.getCell(0).setCellValue(data.getUid());
                            destRow.getCell(1).setCellValue(data.getName());
                            destRow.getCell(2).setCellValue(data.getCardNo());
                            destRow.getCell(3).setCellValue(data.getBalance());
                        });

                        return sheetExportConfig;
                    });
                    return exportConfig;
                });
        return ApiResponse.success("请求成功", taskId);
    }

}
