package com.example.coeusproxy;

import org.springframework.boot.context.properties.ConfigurationProperties;

@ConfigurationProperties(prefix = "coeus")
public record CoeusProperties(
        String baseUrl,
        String apiKey,
        String authType,
        String endUserToken
) {
}
