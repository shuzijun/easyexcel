package com.alibaba.excel.metadata;

import com.alibaba.excel.converters.Converter;

/**
 * the transformation class bound to head
 *
 * @author shuzijun
 * @Date 2020-01-17 14:01
 */
public class HeadConverter {

    /**
     * Column index of head
     */
    private Integer columnIndex;
    /**
     * It only has values when passed in {@link Sheet#setClazz(Class)} and {@link Table#setClazz(Class)}
     */
    private String fieldName;

    /**
     * converter
     */
    Converter converter;

    public Integer getColumnIndex() {
        return columnIndex;
    }

    public void setColumnIndex(Integer columnIndex) {
        this.columnIndex = columnIndex;
    }

    public String getFieldName() {
        return fieldName;
    }

    public void setFieldName(String fieldName) {
        this.fieldName = fieldName;
    }

    public Converter getConverter() {
        return converter;
    }

    public void setConverter(Converter converter) {
        this.converter = converter;
    }
}
