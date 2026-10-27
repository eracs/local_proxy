package com.example.coeusproxy;

import org.springframework.context.annotation.Bean;
import org.springframework.context.annotation.Configuration;
import org.springframework.web.reactive.function.client.WebClient;

@Configuration
public class CoeusWebClientConfig {

    @Bean
    WebClient coeusWebClient(WebClient.Builder builder, CoeusProperties properties) {
        return builder
                .baseUrl(properties.baseUrl().replaceAll("/+$", ""))
                .defaultHeaders(headers -> {
                    headers.setBearerAuth(properties.apiKey());
                    headers.set("AuthType", properties.authType());
                    if (properties.endUserToken() != null && !properties.endUserToken().isBlank()) {
                        headers.set("EndUserToken", properties.endUserToken());
                    }
                })
                .build();
    }
}
