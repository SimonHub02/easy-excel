package com.zb.read;

import com.alibaba.excel.context.AnalysisContext;
import com.alibaba.excel.event.AnalysisEventListener;
import com.alibaba.excel.metadata.CellExtra;
import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

import java.util.ArrayList;
import java.util.List;

/**
 * @param <T> 约定
 * @author shuos
 * 读取监听器
 */
@Data
@AllArgsConstructor
@NoArgsConstructor
public class EasyExcelListener<T> extends AnalysisEventListener<T> {

    private List<T> datas;

    private Integer rowIndex;

    private List<CellExtra> extraMergeInfoList;

    public EasyExcelListener(Integer rowIndex) {
        this.rowIndex = rowIndex;
        datas = new ArrayList<>();
        extraMergeInfoList = new ArrayList<>();
    }

    @Override
    public void invoke(T data, AnalysisContext context) {
        // 是否忽略空行数据，因为自己要做数据校验，所以还是加上，可以根据业务情况使用
        context.readWorkbookHolder().setIgnoreEmptyRow(false);
        datas.add(data);
    }

    @Override
    public void doAfterAllAnalysed(AnalysisContext context) {
    }

    @Override
    public void extra(CellExtra extra, AnalysisContext context) {
        switch (extra.getType()) {
            //是合并的单元格存起来处理
            case MERGE:
                if (extra.getRowIndex() >= rowIndex) {
                    extraMergeInfoList.add(extra);
                }
                break;
            default:
        }
    }


    public List<T> getData() {
        return datas;
    }


    public List<CellExtra> getExtraMergeInfoList() {
        return extraMergeInfoList;
    }

}
