package com.example.demo.vo;

import cn.afterturn.easypoi.excel.annotation.Excel;

public class StudentVO {

    @Excel(name = "年龄")
    private Integer age;

    @Excel(name = "姓名")
    private String name;

    @Excel(name = "手机")
    private String phone;

    public StudentVO() {
    }

    public StudentVO(Integer age, String name, String phone) {
        this.age = age;
        this.name = name;
        this.phone = phone;
    }

    public Integer getAge() {
        return age;
    }

    public void setAge(Integer age) {
        this.age = age;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public String getPhone() {
        return phone;
    }

    public void setPhone(String phone) {
        this.phone = phone;
    }

    @Override
    public String toString() {
        return "StudentVO{" +
                "age=" + age +
                ", name='" + name + '\'' +
                ", phone='" + phone + '\'' +
                '}';
    }
}
