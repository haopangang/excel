package hpg.test;

import com.alibaba.excel.ExcelReader;
import com.alibaba.excel.context.AnalysisContext;
import com.alibaba.excel.event.AnalysisEventListener;
import com.alibaba.excel.metadata.Sheet;
import com.alibaba.excel.support.ExcelTypeEnum;
import function.model.LoanInfo;
import org.junit.Test;

import java.io.*;
import java.lang.reflect.ParameterizedType;
import java.lang.reflect.Type;
import java.util.ArrayList;
import java.util.List;

public class ReadExcelTest {

    @Test
    public void readExcel() throws FileNotFoundException, InterruptedException {
        List<List<String>> all = new ArrayList<>();
        InputStream inputStream = new FileInputStream(new File("E:\\easyExcelDemo.xlsx"));
        try {
            ExcelReader reader = new ExcelReader(inputStream, ExcelTypeEnum.XLSX, null,
                    new AnalysisEventListener<List<String>>() {
                        @Override
                        public void invoke(List<String> object, AnalysisContext context) {
                            int sheetNo = context.getCurrentSheet().getSheetNo();
                            Integer currentRowNum = context.getCurrentRowNum();
                            Sheet currentSheet = context.getCurrentSheet();
//                            new Thread(()->{
//                                System.out.println(
//                                        "当前sheet:" + sheetNo + " 当前行：" +currentRowNum
//                                                + " data:" + object);
//                            }).start();

                            all.add(object);
                        }
                        @Override
                        public void doAfterAllAnalysed(AnalysisContext context) {
                        System.out.println("读取Excel完成");
                        }
                    });

            reader.read();
            Thread.sleep(20000);
            System.out.println("读取结束");
        } catch (Exception e) {
            e.printStackTrace();

        } finally {
            try {
                inputStream.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }


}


