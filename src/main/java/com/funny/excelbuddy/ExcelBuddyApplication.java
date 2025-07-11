package com.funny.excelbuddy;

import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.springframework.boot.autoconfigure.jdbc.DataSourceAutoConfiguration;
import org.springframework.web.servlet.config.annotation.EnableWebMvc;

@EnableWebMvc
@SpringBootApplication(scanBasePackages = {"com.funny"}, exclude = {DataSourceAutoConfiguration.class})
public class ExcelBuddyApplication {

    public static void main(String[] args) {
        SpringApplication.run(ExcelBuddyApplication.class, args);
    }

}
