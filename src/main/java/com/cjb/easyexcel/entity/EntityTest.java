package com.cjb.easyexcel.entity;

import com.github.crab2died.annotation.ExcelField;
import lombok.Data;


@Data
public class EntityTest {


    @ExcelField(title = "第一个字段", order = 1)
    private String field1;

    @ExcelField(title = "第二个字段", order = 2)
    private String field2;

}
