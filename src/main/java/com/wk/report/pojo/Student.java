package com.wk.report.pojo;

import lombok.Data;

import java.util.Date;

@Data
public class Student {

    /**学生id*/
    private int id;
    /**学生姓名*/
    private String name;
    /**学生性别 1：男 2：女*/
    private int sex;
    /**学生年龄*/
    private int age;
    /**学生学号*/
    private int student_no;
    /**学生出生年月*/
    private String birthday;
    /**学生创建时间*/
    private Date create_time;
}
