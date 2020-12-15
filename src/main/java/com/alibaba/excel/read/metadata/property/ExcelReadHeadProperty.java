package com.alibaba.excel.read.metadata.property;

import com.alibaba.excel.metadata.HeadConverter;
import com.alibaba.excel.metadata.Holder;
import com.alibaba.excel.metadata.property.ExcelHeadProperty;

import java.util.List;

/**
 * Define the header attribute of excel
 *
 * @author jipengfei
 */
public class ExcelReadHeadProperty extends ExcelHeadProperty {

    public ExcelReadHeadProperty(Holder holder, Class headClazz, List<List<String>> head, List<HeadConverter> headConverter, Boolean convertAllFiled) {
        super(holder, headClazz, head, headConverter, convertAllFiled);
    }
}
