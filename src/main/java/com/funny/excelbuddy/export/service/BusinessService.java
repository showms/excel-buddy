package com.funny.excelbuddy.export.service;

import com.github.pagehelper.PageInfo;
import com.google.common.collect.ImmutableList;
import com.funny.excelbuddy.export.dto.BusinessExportDataDTO;
import com.funny.excelbuddy.export.req.GetDataPageRequest;
import org.springframework.stereotype.Service;

import java.util.List;

/**
 * @author: 魏云飞 (weiyunfei@rd.keytop.com.cn)
 * @desc:
 * @date: 2025/2/27
 */
@Service
public class BusinessService {

    public PageInfo<BusinessExportDataDTO> listData(GetDataPageRequest pageRequest){
        BusinessExportDataDTO data1 = BusinessExportDataDTO.Builder.create().withUid("1000").withName("飞哥").withCardNo("M10086").withBalance(2800).build();
        BusinessExportDataDTO data2 = BusinessExportDataDTO.Builder.create().withUid("1000").withName("飞哥").withCardNo("M10087").withBalance(9100).build();
        BusinessExportDataDTO data3 = BusinessExportDataDTO.Builder.create().withUid("1000").withName("飞哥").withCardNo("M10088").withBalance(8888).build();
        BusinessExportDataDTO data4 = BusinessExportDataDTO.Builder.create().withUid("1001").withName("小明").withCardNo("M10001").withBalance(4180).build();
        BusinessExportDataDTO data5 = BusinessExportDataDTO.Builder.create().withUid("1002").withName("阿红").withCardNo("M12979").withBalance(6662).build();
        BusinessExportDataDTO data6 = BusinessExportDataDTO.Builder.create().withUid("1002").withName("阿红").withCardNo("M13459").withBalance(2690).build();
        BusinessExportDataDTO data7 = BusinessExportDataDTO.Builder.create().withUid("1003").withName("小白").withCardNo("M14898").withBalance(4866).build();
        List<BusinessExportDataDTO> list = ImmutableList.of(data1, data2, data3, data4, data5, data6, data7);
        return new PageInfo<>(list);
    }

    public PageInfo<BusinessExportDataDTO> listExportData(GetDataPageRequest pageRequest){
        return this.listData(pageRequest);
    }
}
