package com.alibaba.excel.metadata;

import com.alibaba.excel.converters.Converter;

import java.util.List;

/**
 * Basic parameter
 *
 * @author Jiaju Zhuang
 **/
public class BasicParameter {
    /**
     * You can only choose one of the {@link BasicParameter#head} and {@link BasicParameter#clazz}
     */
    private List<List<String>> head;
    /**
     * You can only choose one of the {@link BasicParameter#head} and {@link BasicParameter#clazz}
     */
    private Class clazz;

    /**
     * the transformation class bound to head
     */
    private List<HeadConverter> headConverter;
    /**
     * Custom type conversions override the default
     */
    private List<Converter> customConverterList;
    /**
     * Automatic trim includes sheet name and content
     */
    private Boolean autoTrim;
    /**
     * true if date uses 1904 windowing, or false if using 1900 date windowing.
     *
     * default is false
     *
     * @return
     */
    private Boolean use1904windowing;

    public List<List<String>> getHead() {
        return head;
    }

    public void setHead(List<List<String>> head) {
        this.head = head;
    }

    public Class getClazz() {
        return clazz;
    }

    public void setClazz(Class clazz) {
        this.clazz = clazz;
    }

    public List<Converter> getCustomConverterList() {
        return customConverterList;
    }

    public void setCustomConverterList(List<Converter> customConverterList) {
        this.customConverterList = customConverterList;
    }

    public Boolean getAutoTrim() {
        return autoTrim;
    }

    public void setAutoTrim(Boolean autoTrim) {
        this.autoTrim = autoTrim;
    }

    public Boolean getUse1904windowing() {
        return use1904windowing;
    }

    public void setUse1904windowing(Boolean use1904windowing) {
        this.use1904windowing = use1904windowing;
    }

    public List<HeadConverter> getHeadConverter() {
        return headConverter;
    }

    public void setHeadConverter(List<HeadConverter> headConverter) {
        this.headConverter = headConverter;
    }
}
