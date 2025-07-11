package com.funny.excelbuddy.export.dto;

import org.springframework.beans.BeanUtils;

import java.io.Serializable;

/**
 * @author chenquanwei（chenquanwei@keytop.com.cn）
 * @date 2021/7/1
 */
public class BaseDTO implements Serializable {

	public Object doForward(Object obj){
		BeanUtils.copyProperties(this,obj);
		return obj;
	}

	public Object doBackward(Object obj){
		BeanUtils.copyProperties(obj,this);
		return this;
	}
}
