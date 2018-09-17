package com.xu.aboutpoi;

import org.mybatis.spring.annotation.MapperScan;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;

@SpringBootApplication
@MapperScan("com.xu.aboutpoi.mybatis")
public class AboutPoiApplication {

    public static void main(String[] args) {
        SpringApplication.run(AboutPoiApplication.class, args);
    }
}
