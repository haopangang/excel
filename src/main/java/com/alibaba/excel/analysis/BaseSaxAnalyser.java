package com.alibaba.excel.analysis;

import com.alibaba.excel.context.AnalysisContext;
import com.alibaba.excel.event.AnalysisEventListener;
import com.alibaba.excel.event.AnalysisEventRegisterCenter;
import com.alibaba.excel.event.OneRowAnalysisFinishEvent;
import com.alibaba.excel.metadata.BaseRowModel;
import com.alibaba.excel.metadata.Sheet;
import com.alibaba.excel.support.ExcelTypeEnum;
import com.alibaba.excel.util.TypeUtil;

import java.io.InputStream;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

/**
 *
 * @author jipengfei
 */
public abstract class BaseSaxAnalyser implements AnalysisEventRegisterCenter, ExcelAnalyser {

    protected AnalysisContext analysisContext;

    private LinkedHashMap<String, AnalysisEventListener> listeners = new LinkedHashMap<String, AnalysisEventListener>();

    /**
     */
    protected abstract void execute();

    public void init(InputStream inputStream, ExcelTypeEnum excelTypeEnum, Object custom,
                     AnalysisEventListener eventListener, boolean trim) {
    }

    public void appendLister(String name, AnalysisEventListener listener) {
        if (!listeners.containsKey(name)) {
            listeners.put(name, listener);
        }
    }

    public void analysis(Sheet sheetParam) {
        execute();
    }

    public void analysis() {
        execute();
    }

    /**
     */
    public void cleanAllListeners() {
        listeners = new LinkedHashMap<String, AnalysisEventListener>();
    }

    /**
     * 执行我们设置的listener时间方法，
     * 无论Excel的类型是什么
     * @date 2018-
     * @updateBy pangang.hao@hand-china.com
     * @param event 事件
     */
    public void notifyListeners(OneRowAnalysisFinishEvent event) {
        analysisContext.setCurrentRowAnalysisResult(event.getData());

        Class<? extends BaseRowModel> clazz = analysisContext.getCurrentSheet().getClazz();
        // 根据当前行是否属于，ExcelHeadProperty
        if (analysisContext.getCurrentRowNum() < analysisContext.getCurrentSheet().getHeadLineMun()) {
            if (analysisContext.getCurrentRowNum() <= analysisContext.getCurrentSheet().getHeadLineMun() - 1) {
                analysisContext.buildExcelHeadProperty(null,
                    (List<String>)analysisContext.getCurrentRowAnalysisResult());
            }
        } else {

            analysisContext.setCurrentRowAnalysisResult(event.getData());
            if (listeners.size() == 1) {
                analysisContext.setCurrentRowAnalysisResult(converter((List<String>)event.getData()));
            }
            for (Map.Entry<String, AnalysisEventListener> entry : listeners.entrySet()) {
                AnalysisEventListener analysisEventListener = entry.getValue();
                Object currentRowAnalysisResult = analysisContext.getCurrentRowAnalysisResult();
                if(analysisEventListener.validate(currentRowAnalysisResult,analysisContext)){
                    analysisEventListener.invoke(currentRowAnalysisResult, analysisContext);
                }else {
                    break;
                }
            }
        }
    }

    private List<String> converter(List<String> data) {
        List<String> list = new ArrayList<String>();
        if (data != null) {
            for (String str : data) {
                list.add(TypeUtil.formatFloat(str));
            }
        }
        return list;
    }

}
