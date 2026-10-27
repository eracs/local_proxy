package com.example.coeusproxy;

import io.netty.handler.ssl.SslContextBuilder;
import io.netty.handler.ssl.util.InsecureTrustManagerFactory;
import reactor.netty.http.client.HttpClient;

import org.springframework.context.annotation.Bean;
import org.springframework.context.annotation.Configuration;
import org.springframework.http.client.reactive.ReactorClientHttpConnector;
import org.springframework.web.reactive.function.client.WebClient;

@Configuration
public class CoeusWebClientConfig {

    @Bean
    WebClient coeusWebClient(WebClient.Builder builder, CoeusProperties properties) {
        WebClient.Builder clientBuilder = builder
                .baseUrl(properties.baseUrl().replaceAll("/+$", ""))
                .defaultHeaders(headers -> {
                    headers.setBearerAuth(properties.apiKey());
                    headers.set("AuthType", properties.authType());
                    if (properties.endUserToken() != null && !properties.endUserToken().isBlank()) {
                        headers.set("EndUserToken", properties.endUserToken());
                    }
                });

        if (!properties.sslVerify()) {
            HttpClient insecureClient = HttpClient.create()
                    .secure(ssl -> ssl.sslContext(SslContextBuilder.forClient()
                            .trustManager(InsecureTrustManagerFactory.INSTANCE)));
            clientBuilder.clientConnector(new ReactorClientHttpConnector(insecureClient));
        }

        return clientBuilder.build();
    }
}
