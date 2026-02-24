package com.example.researchtool.util;

import org.springframework.web.client.RestTemplate;
import org.springframework.http.*;

import java.util.*;

public class OpenAIClient {

    private static final String API_KEY = System.getenv("OPENAI_API_KEY");

    public static String normalize(String label) {

        try {
            RestTemplate restTemplate = new RestTemplate();

            String url = "https://api.openai.com/v1/chat/completions";

            Map<String, Object> body = new HashMap<>();
            body.put("model", "gpt-4o-mini");

            List<Map<String, String>> messages = new ArrayList<>();
            messages.add(Map.of("role", "user",
                    "content", "Map this financial line item to standard category: " + label));
            body.put("messages", messages);
            body.put("temperature", 0);

            HttpHeaders headers = new HttpHeaders();
            headers.setBearerAuth(API_KEY);
            headers.setContentType(MediaType.APPLICATION_JSON);

            HttpEntity<Map<String, Object>> request = new HttpEntity<>(body, headers);

            ResponseEntity<Map> response = restTemplate.postForEntity(url, request, Map.class);

            List choices = (List) response.getBody().get("choices");
            Map first = (Map) choices.get(0);
            Map message = (Map) first.get("message");

            return message.get("content").toString().trim();

        } catch (Exception e) {
            return "Unmapped";
        }
    }
}