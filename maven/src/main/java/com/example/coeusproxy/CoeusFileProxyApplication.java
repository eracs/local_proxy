package com.example.coeusproxy;

import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.springframework.boot.context.properties.EnableConfigurationProperties;

@SpringBootApplication
@EnableConfigurationProperties(CoeusProperties.class)
public class CoeusFileProxyApplication {

    public static void main(String[] args) {
        SpringApplication.run(CoeusFileProxyApplication.class, args);
    }
}
