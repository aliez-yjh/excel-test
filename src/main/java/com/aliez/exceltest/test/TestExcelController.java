package com.aliez.exceltest.test;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.ExcelReader;
import com.alibaba.excel.ExcelWriter;
import com.alibaba.excel.read.metadata.ReadSheet;
import com.alibaba.excel.write.metadata.WriteSheet;
import com.aliez.exceltest.entity.FirstSheetVO;
import com.aliez.exceltest.entity.SecondSheetVO;
import lombok.extern.slf4j.Slf4j;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.ResponseBody;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

/**
 * @ClassName TestExcelController
 * @Description TODO
 * @Author aliez
 * @Date 2022/8/28 22:05
 * @Version 1.0
 **/
@Controller
@Slf4j
public class TestExcelController {

    /**
     * easyExcel 导出
     * @param response
     */
    @GetMapping("/export")
    public void exportExcel(HttpServletResponse response){
        String file_name = null;
        try {
            file_name = new String("dicom影像描述匹配模板1.0".getBytes(), "ISO-8859-1");
            response.setContentType("application/vnd.ms-excel");
            response.setHeader("Content-Disposition","attachment;filename="+file_name+".xlsx");
            List<FirstSheetVO> firstSheetVOS = new ArrayList<>();
            FirstSheetVO firstSheetVO = new FirstSheetVO();
            firstSheetVO.setOrderNum(1);
            firstSheetVO.setAge(24);
            firstSheetVO.setBirthDate(new Date());
            firstSheetVO.setGender("男");
            firstSheetVO.setProfessional("教师");
            firstSheetVO.setName("李华");
            firstSheetVOS.add(firstSheetVO);
            List<SecondSheetVO> secondSheetVOS = new ArrayList<>();
            SecondSheetVO secondSheetVO = new SecondSheetVO();
            secondSheetVO.setOrderNum(1);
            secondSheetVO.setEducation("本科");
            secondSheetVO.setWorkplace("广东");
            secondSheetVOS.add(secondSheetVO);
            // 表一写入
            ExcelWriter writer = EasyExcel.write(response.getOutputStream(), FirstSheetVO.class).build();
            WriteSheet sheet = EasyExcel.writerSheet(0, "基础信息").build();
            writer.write(firstSheetVOS,sheet);

            // 表二写入
            WriteSheet sheet2 = EasyExcel.writerSheet(1, "详细信息").head(SecondSheetVO.class).build();
            writer.write(secondSheetVOS,sheet2);

            // 关闭流
            writer.finish();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    /**
     * easyExcel 导入
     * @param file
     * @return
     */
    @PostMapping("/upload")
    @ResponseBody
    public String upload(MultipartFile file){
        try {
            ExcelReader reader = EasyExcel.read(file.getInputStream()).build();
            List<ReadSheet> sheets = reader.excelExecutor().sheetList();
            System.out.println("上传成功。。。");
            for (int i = 0; i < sheets.size(); i++) {
                ReadSheet readSheet = sheets.get(i);
                System.out.println("表序号：" + readSheet.getSheetNo());
                System.out.println("表名：" + readSheet.getSheetName());
                List<Object> objects = EasyExcel.read(file.getInputStream()).sheet(i).doReadSync();
                objects.forEach(System.out::println);
            }
        } catch (IOException e) {
            log.info(e.getMessage(),e);
            return "fail";
        }
        return "success";
    }
}
