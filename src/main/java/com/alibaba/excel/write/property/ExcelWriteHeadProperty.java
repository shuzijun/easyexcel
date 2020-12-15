package com.alibaba.excel.write.property;

import com.alibaba.excel.annotation.format.NumberFormat;
import com.alibaba.excel.annotation.write.style.*;
import com.alibaba.excel.converters.ConverterKeyBuild;
import com.alibaba.excel.converters.DefaultConverterLoader;
import com.alibaba.excel.enums.CellDataTypeEnum;
import com.alibaba.excel.enums.HeadKindEnum;
import com.alibaba.excel.metadata.CellRange;
import com.alibaba.excel.metadata.Head;
import com.alibaba.excel.metadata.HeadConverter;
import com.alibaba.excel.metadata.Holder;
import com.alibaba.excel.metadata.property.*;

import java.lang.reflect.Field;
import java.util.*;

/**
 * Define the header attribute of excel
 *
 * @author jipengfei
 */
public class ExcelWriteHeadProperty extends ExcelHeadProperty {

    private RowHeightProperty headRowHeightProperty;
    private RowHeightProperty contentRowHeightProperty;
    private OnceAbsoluteMergeProperty onceAbsoluteMergeProperty;

    public ExcelWriteHeadProperty(Holder holder, Class headClazz, List<List<String>> head, List<HeadConverter> headConverter, Boolean convertAllFiled) {
        super(holder, headClazz, head, headConverter, convertAllFiled);
        if (getHeadKind() != HeadKindEnum.CLASS) {
            return;
        }
        this.headRowHeightProperty =
            RowHeightProperty.build((HeadRowHeight) headClazz.getAnnotation(HeadRowHeight.class));
        this.contentRowHeightProperty =
            RowHeightProperty.build((ContentRowHeight) headClazz.getAnnotation(ContentRowHeight.class));
        this.onceAbsoluteMergeProperty =
            OnceAbsoluteMergeProperty.build((OnceAbsoluteMerge) headClazz.getAnnotation(OnceAbsoluteMerge.class));

        ColumnWidth parentColumnWidth = (ColumnWidth) headClazz.getAnnotation(ColumnWidth.class);
        HeadStyle parentHeadStyle = (HeadStyle) headClazz.getAnnotation(HeadStyle.class);
        HeadFontStyle parentHeadFontStyle = (HeadFontStyle) headClazz.getAnnotation(HeadFontStyle.class);
        ContentStyle parentContentStyle = (ContentStyle) headClazz.getAnnotation(ContentStyle.class);
        ContentFontStyle parentContentFontStyle = (ContentFontStyle) headClazz.getAnnotation(ContentFontStyle.class);

        for (Map.Entry<Integer, ExcelContentProperty> entry : getContentPropertyMap().entrySet()) {
            Integer index = entry.getKey();
            ExcelContentProperty excelContentPropertyData = entry.getValue();
            if (excelContentPropertyData == null) {
                throw new IllegalArgumentException(
                    "Passing in the class and list the head, the two must be the same size.");
            }
            Field field = excelContentPropertyData.getField();
            Head headData = getHeadMap().get(index);
            ColumnWidth columnWidth = field.getAnnotation(ColumnWidth.class);
            if (columnWidth == null) {
                columnWidth = parentColumnWidth;
            }
            headData.setColumnWidthProperty(ColumnWidthProperty.build(columnWidth));

            HeadStyle headStyle = field.getAnnotation(HeadStyle.class);
            if (headStyle == null) {
                headStyle = parentHeadStyle;
            }
            headData.setHeadStyleProperty(StyleProperty.build(headStyle));

            HeadFontStyle headFontStyle = field.getAnnotation(HeadFontStyle.class);
            if (headFontStyle == null) {
                headFontStyle = parentHeadFontStyle;
            }
            headData.setHeadFontProperty(FontProperty.build(headFontStyle));

            ContentStyle contentStyle = field.getAnnotation(ContentStyle.class);
            if (contentStyle == null) {
                contentStyle = parentContentStyle;
            }
            headData.setContentStyleProperty(StyleProperty.build(contentStyle));

            ContentFontStyle contentFontStyle = field.getAnnotation(ContentFontStyle.class);
            if (contentFontStyle == null) {
                contentFontStyle = parentContentFontStyle;
            }
            headData.setContentFontProperty(FontProperty.build(contentFontStyle));

            headData.setLoopMergeProperty(LoopMergeProperty.build(field.getAnnotation(ContentLoopMerge.class)));
            // If have @NumberFormat, 'NumberStringConverter' is specified by default
            if (excelContentPropertyData.getConverter() == null) {
                NumberFormat numberFormat = field.getAnnotation(NumberFormat.class);
                if (numberFormat != null) {
                    excelContentPropertyData.setConverter(DefaultConverterLoader.loadAllConverter()
                        .get(ConverterKeyBuild.buildKey(field.getType(), CellDataTypeEnum.STRING)));
                }
            }
        }
    }

    public RowHeightProperty getHeadRowHeightProperty() {
        return headRowHeightProperty;
    }

    public void setHeadRowHeightProperty(RowHeightProperty headRowHeightProperty) {
        this.headRowHeightProperty = headRowHeightProperty;
    }

    public RowHeightProperty getContentRowHeightProperty() {
        return contentRowHeightProperty;
    }

    public void setContentRowHeightProperty(RowHeightProperty contentRowHeightProperty) {
        this.contentRowHeightProperty = contentRowHeightProperty;
    }

    public OnceAbsoluteMergeProperty getOnceAbsoluteMergeProperty() {
        return onceAbsoluteMergeProperty;
    }

    public void setOnceAbsoluteMergeProperty(OnceAbsoluteMergeProperty onceAbsoluteMergeProperty) {
        this.onceAbsoluteMergeProperty = onceAbsoluteMergeProperty;
    }

    /**
     * Calculate all cells that need to be merged
     *
     * @return cells that need to be merged
     */
    public List<CellRange> headCellRangeList() {
        List<CellRange> cellRangeList = new ArrayList<CellRange>();
        Set<String> alreadyRangeSet = new HashSet<String>();
        List<Head> headList = new ArrayList<Head>(getHeadMap().values());
        for (int i = 0; i < headList.size(); i++) {
            Head head = headList.get(i);
            List<String> headNameList = head.getHeadNameList();
            for (int j = 0; j < headNameList.size(); j++) {
                if (alreadyRangeSet.contains(i + "-" + j)) {
                    continue;
                }
                alreadyRangeSet.add(i + "-" + j);
                String headName = headNameList.get(j);
                int lastCol = i;
                int lastRow = j;
                for (int k = i + 1; k < headList.size(); k++) {
                    if (headList.get(k).getHeadNameList().get(j).equals(headName)) {
                        alreadyRangeSet.add(k + "-" + j);
                        lastCol = k;
                    } else {
                        break;
                    }
                }
                Set<String> tempAlreadyRangeSet = new HashSet<String>();
                outer:
                for (int k = j + 1; k < headNameList.size(); k++) {
                    for (int l = i; l <= lastCol; l++) {
                        if (headList.get(l).getHeadNameList().get(k).equals(headName)) {
                            tempAlreadyRangeSet.add(l + "-" + k);
                        } else {
                            break outer;
                        }
                    }
                    lastRow = k;
                    alreadyRangeSet.addAll(tempAlreadyRangeSet);
                }
                if (j == lastRow && i == lastCol) {
                    continue;
                }
                cellRangeList
                    .add(new CellRange(j, lastRow, head.getColumnIndex(), headList.get(lastCol).getColumnIndex()));
            }
        }
        return cellRangeList;
    }
}
