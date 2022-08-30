package com.aliez.exceltest.entity;
import com.alibaba.excel.annotation.ExcelProperty;
import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

import java.util.Date;

/**
 * 演示表-个人基础信息
 */
@Data
@NoArgsConstructor
@AllArgsConstructor
public class FirstSheetVO {

    @ExcelProperty(value = "序号",index = 0)
    private Integer orderNum;
    @ExcelProperty(value = "姓名",index = 1)
    private String name;
    @ExcelProperty(value = "年龄",index = 2)
    private Integer age;
    @ExcelProperty(value = "性别",index = 3)
    private String gender;
    @ExcelProperty(value = "职业",index = 4)
    private String professional;
    @ExcelProperty(value = "出生日期")
    private Date birthDate;
}
