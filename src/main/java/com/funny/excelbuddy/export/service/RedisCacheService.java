package com.funny.excelbuddy.export.service;

import com.alibaba.fastjson.JSON;
import org.springframework.stereotype.Service;
import org.springframework.util.StringUtils;

import java.util.Optional;
import java.util.concurrent.TimeUnit;

/**
 * @author: 魏云飞 (weiyunfei@rd.keytop.com.cn)
 * @desc:
 * @date: 2025/7/11
 */
@Service
public class RedisCacheService {

    public <T> Optional<T> get(String key, Class<T> clazz) {
        Optional<T> rtn = Optional.empty();

        // 读取redis缓存

        return rtn;
    }

    public void setEx(String key, String value, long timeout, TimeUnit unit) {
        // redis setEx
    }
}
