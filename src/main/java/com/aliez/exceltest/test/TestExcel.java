package com.aliez.exceltest.test;

import cn.afterturn.easypoi.excel.ExcelExportUtil;
import cn.afterturn.easypoi.excel.ExcelImportUtil;
import cn.afterturn.easypoi.excel.entity.ExportParams;
import cn.afterturn.easypoi.excel.entity.ImportParams;
import com.aliez.exceltest.entity.UserEntity;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.File;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.Map;

/**
 * @ClassName TestExcel
 * @Description Excel 导入导出测试
 * @Author aliez
 * @Date 2022/8/28 21:47
 * @Version 1.0
 **/
public class TestExcel {

    public static void main(String[] args) throws Exception {
        testEasyPoiRead();
    }

    /**
     * 测试 EasyPoi 导出功能
     * @throws Exception
     */
    public static void testEasyPoiWrite() throws Exception{
        List<UserEntity> dataList = new ArrayList<>();
        for (int i = 0; i < 10; i++) {
            UserEntity userEntity = new UserEntity();
            userEntity.setName("李四" + i);
            userEntity.setAge(30 + i);
            userEntity.setTime(new Date(System.currentTimeMillis() + i));
            dataList.add(userEntity);
        }
        //生成excel文档
        Workbook workbook = ExcelExportUtil.exportExcel(new ExportParams("用户","用户信息"),
                UserEntity.class, dataList);
        FileOutputStream fos = new FileOutputStream("D:/easypoi-user1.xls");
        workbook.write(fos);
        fos.close();
    }

    /**
     * 测试 EasyPoi 导入功能
     * @throws Exception
     */
    public static void testEasyPoiRead() throws Exception{
        ImportParams params = new ImportParams();
        params.setTitleRows(1);
        params.setHeadRows(1);
        long start = System.currentTimeMillis();
        List<Map<String, Object>> list = ExcelImportUtil.importExcel(new File("D:/easypoi-user1.xls"),
                Map.class, params);
        System.out.println(System.currentTimeMillis() - start);
        System.out.println(list);
    }

}
