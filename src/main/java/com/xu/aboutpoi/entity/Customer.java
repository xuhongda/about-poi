package com.xu.aboutpoi.entity;

import lombok.Data;
import lombok.ToString;

/**
 * @author xuhongda on 2018/9/17
 * com.xu.aboutpoi.entity
 * about-poi
 */
@Data
@ToString
public class Customer {
    private Long id;
    private String name;
    private String contact;
    private String email;
    private String telephone;
    private String remark;
}
