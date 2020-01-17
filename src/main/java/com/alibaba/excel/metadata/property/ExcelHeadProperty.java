package com.alibaba.excel.metadata.property;

import com.alibaba.excel.annotation.ExcelProperty;
import com.alibaba.excel.annotation.format.DateTimeFormat;
import com.alibaba.excel.annotation.format.NumberFormat;
import com.alibaba.excel.converters.AutoConverter;
import com.alibaba.excel.converters.Converter;
import com.alibaba.excel.enums.HeadKindEnum;
import com.alibaba.excel.exception.ExcelCommonException;
import com.alibaba.excel.metadata.Head;
import com.alibaba.excel.metadata.HeadConverter;
import com.alibaba.excel.metadata.Holder;
import com.alibaba.excel.util.ClassUtils;
import com.alibaba.excel.util.CollectionUtils;
import com.alibaba.excel.util.StringUtils;
import com.alibaba.excel.write.metadata.holder.AbstractWriteHolder;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.lang.reflect.Field;
import java.util.*;

/**
 * Define the header attribute of excel
 *
 * @author jipengfei
 */
public class ExcelHeadProperty {

    private static final Logger LOGGER = LoggerFactory.getLogger(ExcelHeadProperty.class);
    /**
     * Custom class
     */
    private Class headClazz;
    /**
     * The types of head
     */
    private HeadKindEnum headKind;
    /**
     * The number of rows in the line with the most rows
     */
    private int headRowNumber;
    /**
     * Configuration header information
     */
    private Map<Integer, Head> headMap;
    /**
     * Configuration column information
     */
    private Map<Integer, ExcelContentProperty> contentPropertyMap;
    /**
     * Configuration column information
     */
    private Map<String, ExcelContentProperty> fieldNameContentPropertyMap;
    /**
     * Fields ignored
     */
    private Map<String, Field> ignoreMap;

    public ExcelHeadProperty(Holder holder, Class headClazz, List<List<String>> head, List<HeadConverter> headConverter, Boolean convertAllFiled) {
        this.headClazz = headClazz;
        headMap = new TreeMap<Integer, Head>();
        contentPropertyMap = new TreeMap<Integer, ExcelContentProperty>();
        fieldNameContentPropertyMap = new HashMap<String, ExcelContentProperty>();
        ignoreMap = new HashMap<String, Field>(16);
        headKind = HeadKindEnum.NONE;
        headRowNumber = 0;
        if (head != null && !head.isEmpty()) {
            Map<Integer,HeadConverter> headConverterMap = new HashMap<Integer, HeadConverter>();
            if(!CollectionUtils.isEmpty(headConverter)){
                for (HeadConverter converter : headConverter) {
                    headConverterMap.put(converter.getColumnIndex(),converter);
                }
            }

            int headIndex = 0;
            for (int i = 0; i < head.size(); i++) {
                if (holder instanceof AbstractWriteHolder) {
                    if (((AbstractWriteHolder)holder).ignore(null, i)) {
                        continue;
                    }
                }
                headMap.put(headIndex, new Head(headIndex, null, head.get(i), Boolean.FALSE, Boolean.TRUE));
                if(headConverterMap.containsKey(headIndex)) {
                    ExcelContentProperty excelContentProperty = new ExcelContentProperty();
                    excelContentProperty.setConverter(headConverterMap.get(headIndex).getConverter());
                    contentPropertyMap.put(headIndex, excelContentProperty);
                } else {
                    contentPropertyMap.put(headIndex, null);
                }
                headIndex++;
            }
            headKind = HeadKindEnum.STRING;
        } else {
            // convert headClazz to head
            initColumnProperties(holder, convertAllFiled, headConverter);
        }
        initHeadRowNumber();
        if (LOGGER.isDebugEnabled()) {
            LOGGER.debug("The initialization sheet/table 'ExcelHeadProperty' is complete , head kind is {}", headKind);
        }
    }

    private void initHeadRowNumber() {
        headRowNumber = 0;
        for (Head head : headMap.values()) {
            List<String> list = head.getHeadNameList();
            if (list != null && list.size() > headRowNumber) {
                headRowNumber = list.size();
            }
        }
        for (Head head : headMap.values()) {
            List<String> list = head.getHeadNameList();
            if (list != null && !list.isEmpty() && list.size() < headRowNumber) {
                int lack = headRowNumber - list.size();
                int last = list.size() - 1;
                for (int i = 0; i < lack; i++) {
                    list.add(list.get(last));
                }
            }
        }
    }

    private void initColumnProperties(Holder holder, Boolean convertAllFiled, List<HeadConverter> headConverter) {
        if (headClazz == null) {
            return;
        }
        // Declared fields
        List<Field> defaultFieldList = new ArrayList<Field>();
        Map<Integer, Field> customFiledMap = new TreeMap<Integer, Field>();
        ClassUtils.declaredFields(headClazz, defaultFieldList, customFiledMap, ignoreMap, convertAllFiled);

        Map<String,HeadConverter> convertHeadMap = new HashMap<String, HeadConverter>();
        if(!CollectionUtils.isEmpty(headConverter)){
            for (HeadConverter converter : headConverter) {
                convertHeadMap.put(converter.getFieldName(),converter);
            }
        }

        int index = 0;
        for (Field field : defaultFieldList) {
            while (customFiledMap.containsKey(index)) {
                Field customFiled = customFiledMap.get(index);
                customFiledMap.remove(index);
                if (!initOneColumnProperty(holder, index, customFiled, Boolean.TRUE, convertHeadMap)) {
                    index++;
                }
            }
            if (!initOneColumnProperty(holder, index, field, Boolean.FALSE, convertHeadMap)) {
                index++;
            }
        }
        for (Map.Entry<Integer, Field> entry : customFiledMap.entrySet()) {
            initOneColumnProperty(holder, entry.getKey(), entry.getValue(), Boolean.TRUE, convertHeadMap);
        }
        headKind = HeadKindEnum.CLASS;
    }

    /**
     * Initialization column property
     *
     * @param holder
     * @param index
     * @param field
     * @param forceIndex
     * @return Ignore current field
     */
    private boolean initOneColumnProperty(Holder holder, int index, Field field, Boolean forceIndex, Map<String,HeadConverter> headConverterMap) {
        if (holder instanceof AbstractWriteHolder) {
            if (((AbstractWriteHolder)holder).ignore(field.getName(), index)) {
                return true;
            }
        }
        ExcelProperty excelProperty = field.getAnnotation(ExcelProperty.class);
        List<String> tmpHeadList = new ArrayList<String>();
        boolean notForceName = excelProperty == null || excelProperty.value().length <= 0
            || (excelProperty.value().length == 1 && StringUtils.isEmpty((excelProperty.value())[0]));
        if (notForceName) {
            tmpHeadList.add(field.getName());
        } else {
            Collections.addAll(tmpHeadList, excelProperty.value());
        }
        Head head = new Head(index, field.getName(), tmpHeadList, forceIndex, !notForceName);
        ExcelContentProperty excelContentProperty = new ExcelContentProperty();
        if (headConverterMap.containsKey(field.getName())) {
            excelContentProperty.setConverter(headConverterMap.get(field.getName()).getConverter());
        } else if (excelProperty != null) {
            Class<? extends Converter> convertClazz = excelProperty.converter();
            if (convertClazz != AutoConverter.class) {
                try {
                    Converter converter = convertClazz.newInstance();
                    excelContentProperty.setConverter(converter);
                } catch (Exception e) {
                    throw new ExcelCommonException("Can not instance custom converter:" + convertClazz.getName());
                }
            }
        }
        excelContentProperty.setHead(head);
        excelContentProperty.setField(field);
        excelContentProperty
            .setDateTimeFormatProperty(DateTimeFormatProperty.build(field.getAnnotation(DateTimeFormat.class)));
        excelContentProperty
            .setNumberFormatProperty(NumberFormatProperty.build(field.getAnnotation(NumberFormat.class)));
        headMap.put(index, head);
        contentPropertyMap.put(index, excelContentProperty);
        fieldNameContentPropertyMap.put(field.getName(), excelContentProperty);
        return false;
    }

    public Class getHeadClazz() {
        return headClazz;
    }

    public void setHeadClazz(Class headClazz) {
        this.headClazz = headClazz;
    }

    public HeadKindEnum getHeadKind() {
        return headKind;
    }

    public void setHeadKind(HeadKindEnum headKind) {
        this.headKind = headKind;
    }

    public boolean hasHead() {
        return headKind != HeadKindEnum.NONE;
    }

    public int getHeadRowNumber() {
        return headRowNumber;
    }

    public void setHeadRowNumber(int headRowNumber) {
        this.headRowNumber = headRowNumber;
    }

    public Map<Integer, Head> getHeadMap() {
        return headMap;
    }

    public void setHeadMap(Map<Integer, Head> headMap) {
        this.headMap = headMap;
    }

    public Map<Integer, ExcelContentProperty> getContentPropertyMap() {
        return contentPropertyMap;
    }

    public void setContentPropertyMap(Map<Integer, ExcelContentProperty> contentPropertyMap) {
        this.contentPropertyMap = contentPropertyMap;
    }

    public Map<String, ExcelContentProperty> getFieldNameContentPropertyMap() {
        return fieldNameContentPropertyMap;
    }

    public void setFieldNameContentPropertyMap(Map<String, ExcelContentProperty> fieldNameContentPropertyMap) {
        this.fieldNameContentPropertyMap = fieldNameContentPropertyMap;
    }

    public Map<String, Field> getIgnoreMap() {
        return ignoreMap;
    }

    public void setIgnoreMap(Map<String, Field> ignoreMap) {
        this.ignoreMap = ignoreMap;
    }
}
