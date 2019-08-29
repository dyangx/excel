package com.example.demo.controller;

import com.example.demo.util.ExcelUtil;
import com.example.demo.vo.StudentVO;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

@RestController
@RequestMapping("excel")
public class ExcelController {

    @RequestMapping("export")
    public void export(HttpServletResponse response) throws IOException {
        List<StudentVO> list = new ArrayList<>();
        StudentVO vo1 = new StudentVO(1,"张三","188");
        StudentVO vo2 = new StudentVO(2,"王麻子","183");
        list.add(vo1);
        list.add(vo2);
        ExcelUtil.exportExcel(list,"studentInfo","shit",StudentVO.class,"excel.xls",response);
    }
    @RequestMapping("export2")
    public void export1(HttpServletResponse response) throws IOException {
        List<StudentVO> list = new ArrayList<>();
        StudentVO vo1 = new StudentVO(1,"张三","188");
        StudentVO vo2 = new StudentVO(2,"王麻子","183");
        list.add(vo1);
        list.add(vo2);
        //要删除的列
        List<String> cols = new ArrayList<>();
        cols.add("年龄");
        ExcelUtil.exportExcel(list,"studentInfo","shit",StudentVO.class,"excel.xls",response,cols);
    }
}
