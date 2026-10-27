package com.example.coeusproxy;

import java.util.OptionalLong;

import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PathVariable;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.reactive.function.client.WebClient;
import org.springframework.web.server.ResponseStatusException;

import reactor.core.publisher.Flux;
import reactor.core.publisher.Mono;

import org.springframework.core.io.buffer.DataBuffer;

@RestController
public class CoeusDocumentController {

    private final WebClient coeusWebClient;

    public CoeusDocumentController(WebClient coeusWebClient) {
        this.coeusWebClient = coeusWebClient;
    }

    @GetMapping("/api/documents/{docId}")
    public Mono<ResponseEntity<Flux<DataBuffer>>> download(@PathVariable String docId) {
        return coeusWebClient
                .get()
                .uri(uriBuilder -> uriBuilder.pathSegment("document", "get", docId).build())
                .exchangeToMono(response -> {
                    if (response.statusCode().isError()) {
                        return response.bodyToMono(String.class)
                                .defaultIfEmpty("")
                                .flatMap(ignored -> Mono.error(new ResponseStatusException(
                                        response.statusCode(), "Coeus download failed")));
                    }

                    HttpHeaders headers = new HttpHeaders();
                    headers.setContentType(response.headers().contentType()
                            .orElse(MediaType.APPLICATION_OCTET_STREAM));
                    copyHeader(response.headers().asHttpHeaders(), headers, HttpHeaders.CONTENT_LENGTH);
                    copyHeader(response.headers().asHttpHeaders(), headers, HttpHeaders.CONTENT_DISPOSITION);

                    return Mono.just(ResponseEntity
                            .status(response.statusCode())
                            .headers(headers)
                            .body(response.bodyToFlux(DataBuffer.class)));
                });
    }

    private static void copyHeader(HttpHeaders source, HttpHeaders target, String name) {
        OptionalLong contentLength = source.getContentLength() >= 0
                ? OptionalLong.of(source.getContentLength())
                : OptionalLong.empty();
        if (HttpHeaders.CONTENT_LENGTH.equalsIgnoreCase(name) && contentLength.isPresent()) {
            target.setContentLength(contentLength.getAsLong());
            return;
        }

        String value = source.getFirst(name);
        if (value != null) {
            target.set(name, value);
        }
    }
}
