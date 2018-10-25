package com.alibaba.excel.modelbuild;

import com.alibaba.excel.context.AnalysisContext;
import com.alibaba.excel.event.AnalysisEventListener;
import com.alibaba.excel.exception.ExcelGenerateException;
import com.alibaba.excel.metadata.CheckoutResult;
import com.alibaba.excel.metadata.ExcelColumnProperty;
import com.alibaba.excel.metadata.ExcelHeadProperty;
import com.alibaba.excel.support.ResultEnum;
import com.alibaba.excel.util.TypeUtil;
import org.apache.commons.beanutils.BeanUtils;

import java.util.List;

/**
 * @author jipengfei
 */
public class ModelBuildEventListener extends AnalysisEventListener {

    @Override
    public boolean validate(Object object, AnalysisContext context) {
        return validator(context, (List<String>) object);
    }

    @Override
    public void invoke(Object object, AnalysisContext context) {
        if (context.getExcelHeadProperty() != null && context.getExcelHeadProperty().getHeadClazz() != null) {
            Object resultModel = buildUserModel(context, (List<String>) object);
            context.setCurrentRowAnalysisResult(resultModel);
        }

    }

    private boolean validator(AnalysisContext context, List<String> stringList) {
        ExcelHeadProperty excelHeadProperty = context.getExcelHeadProperty();
        CheckoutResult checkoutResult = context.getCheckoutResult();
        if (excelHeadProperty != null) {
            Integer currentRowNum = context.getCurrentRowNum();
            if (checkoutResult == null) {
                checkoutResult = CheckoutResult.ok();
                context.setCheckoutResult(checkoutResult);
            }
            StringBuilder rowErrorMsg = new StringBuilder();
            for (int i = 0; i < stringList.size(); i++) {
                ExcelColumnProperty columnProperty = excelHeadProperty.getExcelColumnProperty(i);
                if (columnProperty != null) {
                    TypeUtil.validator(stringList.get(i), columnProperty.getField(),
                            columnProperty.getFormat(), rowErrorMsg);
                }
            }
            if (rowErrorMsg.length() > 0) {
                if (checkoutResult.getStatus() == ResultEnum.OK) {
                    checkoutResult.setStatus(ResultEnum.ERROR);
                }
                checkoutResult.setErrMsg(currentRowNum, rowErrorMsg);
                return false;
            }

        }
        return true;
    }


    private Object buildUserModel(AnalysisContext context, List<String> stringList) {
        ExcelHeadProperty excelHeadProperty = context.getExcelHeadProperty();

        Object resultModel;
        try {
            resultModel = excelHeadProperty.getHeadClazz().newInstance();
        } catch (Exception e) {
            throw new ExcelGenerateException(e);
        }
        if (excelHeadProperty != null) {
            for (int i = 0; i < stringList.size(); i++) {
                ExcelColumnProperty columnProperty = excelHeadProperty.getExcelColumnProperty(i);
                if (columnProperty != null) {
                    Object value = TypeUtil.convert(stringList.get(i), columnProperty.getField(),
                            columnProperty.getFormat(), context.use1904WindowDate());
                    if (value != null) {
                        try {
                            BeanUtils.setProperty(resultModel, columnProperty.getField().getName(), value);
                        } catch (Exception e) {
                            throw new ExcelGenerateException(
                                    columnProperty.getField().getName() + " can not set value " + value, e);
                        }
                    }
                }
            }
        }
        return resultModel;
    }

    @Override
    public void doAfterAllAnalysed(AnalysisContext context) {

    }
}
