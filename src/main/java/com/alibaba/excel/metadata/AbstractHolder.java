package com.alibaba.excel.metadata;

import com.alibaba.excel.converters.Converter;

import java.util.List;
import java.util.Map;

/**
 * Write/read holder
 *
 * @author Jiaju Zhuang
 */
public abstract class AbstractHolder implements ConfigurationHolder {
    /**
     * Record whether it's new or from cache
     */
    private Boolean newInitialization;
    /**
     * You can only choose one of the {@link AbstractHolder#head} and {@link AbstractHolder#clazz}
     */
    private List<List<String>> head;
    /**
     * You can only choose one of the {@link AbstractHolder#head} and {@link AbstractHolder#clazz}
     */
    private Class clazz;
    /**
     * the transformation class bound to head
     */
    private List<HeadConverter> headConverter;
    /**
     * Some global variables
     */
    private GlobalConfiguration globalConfiguration;

    /**
     * <p>
     * Read key:
     * <p>
     * Write key:
     */
    private Map<String, Converter> converterMap;

    public AbstractHolder(BasicParameter basicParameter, AbstractHolder prentAbstractHolder) {
        this.newInitialization = Boolean.TRUE;
        if (basicParameter.getHead() == null && basicParameter.getClazz() == null && prentAbstractHolder != null) {
            this.head = prentAbstractHolder.getHead();
        } else {
            this.head = basicParameter.getHead();
        }
        if (basicParameter.getHead() == null && basicParameter.getClazz() == null && prentAbstractHolder != null) {
            this.clazz = prentAbstractHolder.getClazz();
        } else {
            this.clazz = basicParameter.getClazz();
        }
        if (basicParameter.getHeadConverter() == null  && prentAbstractHolder != null) {
            this.headConverter = prentAbstractHolder.getHeadConverter();
        } else {
            this.headConverter = basicParameter.getHeadConverter();
        }
        this.globalConfiguration = new GlobalConfiguration();
        if (basicParameter.getAutoTrim() == null) {
            if (prentAbstractHolder == null) {
                globalConfiguration.setAutoTrim(Boolean.TRUE);
            } else {
                globalConfiguration.setAutoTrim(prentAbstractHolder.getGlobalConfiguration().getAutoTrim());
            }
        } else {
            globalConfiguration.setAutoTrim(basicParameter.getAutoTrim());
        }
    }

    public Boolean getNewInitialization() {
        return newInitialization;
    }

    public void setNewInitialization(Boolean newInitialization) {
        this.newInitialization = newInitialization;
    }

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

    public GlobalConfiguration getGlobalConfiguration() {
        return globalConfiguration;
    }

    public void setGlobalConfiguration(GlobalConfiguration globalConfiguration) {
        this.globalConfiguration = globalConfiguration;
    }

    public Map<String, Converter> getConverterMap() {
        return converterMap;
    }

    public void setConverterMap(Map<String, Converter> converterMap) {
        this.converterMap = converterMap;
    }

    @Override
    public Map<String, Converter> converterMap() {
        return getConverterMap();
    }

    @Override
    public GlobalConfiguration globalConfiguration() {
        return getGlobalConfiguration();
    }

    @Override
    public boolean isNew() {
        return getNewInitialization();
    }

    public List<HeadConverter> getHeadConverter() {
        return headConverter;
    }

    public void setHeadConverter(List<HeadConverter> headConverter) {
        this.headConverter = headConverter;
    }
}
