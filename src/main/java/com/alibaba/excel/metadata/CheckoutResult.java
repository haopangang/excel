package com.alibaba.excel.metadata;

import com.alibaba.excel.support.ResultEnum;
import org.apache.commons.codec.binary.StringUtils;

import java.util.HashMap;

/**
 * @author pangang.hao@hand-china.com
 * @version 1.0
 * @name CheckoutResult
 * @description 校验数据的结果
 * @date 10/24/2018
 */
public class CheckoutResult {
    private HashMap<Integer, StringBuilder> errMsg = new HashMap<>();
    private ResultEnum status;

    public CheckoutResult() {
    }

    public CheckoutResult(ResultEnum status) {
        this.status = status;
    }

    public static CheckoutResult ok(){
        return new CheckoutResult(ResultEnum.OK);
    }

    public static CheckoutResult error(Integer rowNum,String msg){
        return new CheckoutResult(ResultEnum.ERROR)
                .setErrMsg(rowNum,new StringBuilder(msg));
    }


    public void setErrMsg(HashMap<Integer, StringBuilder> errMsg) {
        this.errMsg = errMsg;
    }
    public void setErrMsg(Integer row,String errMsg) {
        setErrMsg(row,new StringBuilder(errMsg));
    }
    public CheckoutResult setErrMsg(Integer row,StringBuilder msg) {
        if(this.errMsg.containsKey(row)){
            StringBuilder stringBuilder = errMsg.get(row);
            stringBuilder.append(msg);
        }else {
            this.errMsg.put(row,msg);
        }
        return this;
    }
    public CheckoutResult setErrMsg(CheckoutResult result) {
        if(result.getStatus()==ResultEnum.ERROR){
            this.setErrMsg(result.getErrMsg());
        }
        return this;
    }

    public HashMap<Integer, StringBuilder> getErrMsg() {
        return errMsg;
    }

    public ResultEnum getStatus() {
        return status;
    }
}
