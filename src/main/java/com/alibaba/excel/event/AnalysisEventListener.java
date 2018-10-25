package com.alibaba.excel.event;

import com.alibaba.excel.context.AnalysisContext;

/**
 *
 *
 * @author jipengfei
 */
public abstract class AnalysisEventListener<T> {

    /**
     *  Check data before analysis
     *  如果校验没有通过，则不会执行invoke方法，
     *  并且不会进入到下一个analysisEventListener中
     * @param object one row data
     * @param context analysis context
     * @return
     */
    public boolean validate(T object, AnalysisContext context){
        return true;
    }

    /**
     * when analysis one row trigger invoke function
     *
     * @param object  one row data
     * @param context analysis context
     */
    public abstract void invoke(T object, AnalysisContext context);

    /**
     * if have something to do after all  analysis
     *
     * @param context
     */
    public abstract void doAfterAllAnalysed(AnalysisContext context);
}
