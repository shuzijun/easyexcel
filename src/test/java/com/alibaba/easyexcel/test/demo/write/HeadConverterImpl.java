package com.alibaba.easyexcel.test.demo.write;

import com.alibaba.excel.converters.Converter;
import com.alibaba.excel.enums.CellDataTypeEnum;
import com.alibaba.excel.metadata.CellData;
import com.alibaba.excel.metadata.GlobalConfiguration;
import com.alibaba.excel.metadata.property.ExcelContentProperty;
import com.alibaba.excel.util.CollectionUtils;

import java.util.Map;

/**
 * Created with IntelliJ IDEA.
 *
 * @author shuzijun
 * @Date 2020-01-17 10:38
 */
public class HeadConverterImpl implements Converter<Object> {

    /**
     * mapping数据
     */
    private Map<String, String> mappings;

    public HeadConverterImpl(Map<String, String> mappings) {
        this.mappings = mappings;
    }

    @Override
    public Class supportJavaTypeKey() {
        return String.class;
    }

    @Override
    public CellDataTypeEnum supportExcelTypeKey() {
        return CellDataTypeEnum.STRING;
    }

    @Override
    public Object convertToJavaData(CellData cellData, ExcelContentProperty contentProperty, GlobalConfiguration globalConfiguration) throws Exception {
        return cellData.getStringValue();
    }

    @Override
    public CellData convertToExcelData(Object value, ExcelContentProperty contentProperty, GlobalConfiguration globalConfiguration) throws Exception {

        if (value == null) {
            return new CellData();
        }

        if (CollectionUtils.isEmpty(mappings)) {
            return new CellData(value.toString());
        } else {
            if (mappings.containsKey(value.toString())) {
                return new CellData(mappings.get(value.toString()));
            } else {
                return new CellData(value.toString());
            }
        }

    }

}
