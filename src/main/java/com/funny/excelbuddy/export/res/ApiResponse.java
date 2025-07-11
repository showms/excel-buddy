package com.funny.excelbuddy.export.res;

/**
 * @author: 魏云飞 (weiyunfei@rd.keytop.com.cn)
 * @desc:
 * @date: 2025/2/27
 */
public class ApiResponse<T> {
    private Integer code;

    private String message;

    private T data;

    public ApiResponse() {
    }

    public ApiResponse(Integer code) {
        this.code = code;
    }

    public ApiResponse(Integer code, String message) {
        this.code = code;
        this.message = message;
    }

    public ApiResponse(Integer code, String message, T data) {
        this.code = code;
        this.message = message;
        this.data = data;
    }

    public Integer getCode() {
        return code;
    }

    public void setCode(Integer code) {
        this.code = code;
    }

    public String getMessage() {
        return message;
    }

    public void setMessage(String message) {
        this.message = message;
    }

    public T getData() {
        return data;
    }

    public void setData(T data) {
        this.data = data;
    }

    public static <T> ApiResponse<T> success() {
        return new ApiResponse<>(2000, "操作成功");
    }

    public static <T> ApiResponse<T> success(String message) {
        return new ApiResponse<>(2000, message);
    }

    public static <T> ApiResponse<T> success(String message, T data) {
        return new ApiResponse<>(2000, message, data);
    }

    public static <T> ApiResponse<T> error() {
        return new ApiResponse<>(5000, "操作失败");
    }

    public static <T> ApiResponse<T> error(String message) {
        return new ApiResponse<>(5000, message);
    }
}
