package com.funny.excelbuddy.export.dto;

/**
 * @author: 魏云飞 (weiyunfei@rd.keytop.com.cn)
 * @desc:
 * @date: 2025/2/27
 */
public class BusinessExportDataDTO {

    private String uid;
    private String name;
    private String cardNo;
    private Integer balance;

    public String getCardNo() {
        return cardNo;
    }

    public void setCardNo(String cardNo) {
        this.cardNo = cardNo;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public Integer getBalance() {
        return balance;
    }

    public void setBalance(Integer balance) {
        this.balance = balance;
    }

    public String getUid() {
        return uid;
    }

    public void setUid(String uid) {
        this.uid = uid;
    }

    public static final class Builder {
        private String uid;
        private String name;
        private String cardNo;
        private Integer balance;

        private Builder() {
        }

        public static Builder create() {
            return new Builder();
        }

        public Builder withUid(String uid) {
            this.uid = uid;
            return this;
        }

        public Builder withCardNo(String cardNo) {
            this.cardNo = cardNo;
            return this;
        }

        public Builder withName(String name) {
            this.name = name;
            return this;
        }

        public Builder withBalance(Integer balance) {
            this.balance = balance;
            return this;
        }

        public BusinessExportDataDTO build() {
            BusinessExportDataDTO businessExportDataDTO = new BusinessExportDataDTO();
            businessExportDataDTO.setUid(uid);
            businessExportDataDTO.setName(name);
            businessExportDataDTO.setCardNo(cardNo);
            businessExportDataDTO.setBalance(balance);
            return businessExportDataDTO;
        }
    }
}
