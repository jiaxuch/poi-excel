package com.jiaxuch.poiexcel.config;

/**
 * @author jiaxuch
 * @data 2020/6/27
 */
public class HeaderConfig {
    private String headerName;
    private String fieldName;

    public HeaderConfig() {
    }

    public HeaderConfig(String headerName, String fieldName) {
        this.headerName = headerName;
        this.fieldName = fieldName;
    }

    public String getHeaderName() {
        return headerName;
    }

    public void setHeaderName(String headerName) {
        this.headerName = headerName;
    }

    public String getFieldName() {
        return fieldName;
    }

    public void setFieldName(String fieldName) {
        this.fieldName = fieldName;
    }

    @Override
    public String toString() {
        return "HeaderConfig{" +
                "headerName='" + headerName + '\'' +
                ", fieldName='" + fieldName + '\'' +
                '}';
    }
}
