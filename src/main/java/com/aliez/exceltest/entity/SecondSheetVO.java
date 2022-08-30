package com.aliez.exceltest.entity;
import com.alibaba.excel.annotation.ExcelProperty;
import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

/**
 * 演示表-个人详细信息
 */
@Data
@NoArgsConstructor
@AllArgsConstructor
public class SecondSheetVO {
    @ExcelProperty(value = "序号",index = 0)
    private Integer orderNum;
    @ExcelProperty(value = "学历",index = 1)
    private String education;
    @ExcelProperty(value = "工作地点",index = 2)
    private String workplace;
}
