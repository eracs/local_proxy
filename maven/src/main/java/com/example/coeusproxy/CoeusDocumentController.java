package com.example.coeusproxy;

import org.springframework.core.io.buffer.DataBuffer;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.server.reactive.ServerHttpResponse;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PathVariable;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.reactive.function.client.WebClient;
import org.springframework.web.server.ResponseStatusException;

import reactor.core.publisher.Mono;

@RestController
public class CoeusDocumentController {

    private final WebClient coeusWebClient;

    public CoeusDocumentController(WebClient coeusWebClient) {
        this.coeusWebClient = coeusWebClient;
    }

    @GetMapping("/api/documents/{docId}")
    public Mono<Void> download(
            @PathVariable String docId,
            ServerHttpResponse serverResponse
    ) {
        return coeusWebClient
                .get()
                .uri(uriBuilder -> uriBuilder.pathSegment("document", "get", docId).build())
                .exchangeToMono(upstream -> {
                    if (upstream.statusCode().isError()) {
                        return upstream.bodyToMono(String.class)
                                .defaultIfEmpty("")
                                .flatMap(ignored -> Mono.error(new ResponseStatusException(
                                        upstream.statusCode(), "Coeus download failed")));
                    }

                    HttpHeaders upstreamHeaders = upstream.headers().asHttpHeaders();
                    serverResponse.setStatusCode(upstream.statusCode());
                    MediaType contentType = upstreamHeaders.getContentType();
                    serverResponse.getHeaders().setContentType(contentType != null
                            ? contentType
                            : MediaType.APPLICATION_OCTET_STREAM);
                    copyHeader(upstreamHeaders, serverResponse.getHeaders(),
                            HttpHeaders.CONTENT_DISPOSITION);

                    return serverResponse.writeWith(upstream.bodyToFlux(DataBuffer.class));
                });
    }

    private static void copyHeader(HttpHeaders source, HttpHeaders target, String name) {
        String value = source.getFirst(name);
        if (value != null) {
            target.set(name, value);
        }
    }
}
