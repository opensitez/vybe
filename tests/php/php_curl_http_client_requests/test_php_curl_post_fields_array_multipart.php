<?php
// vybe-test: php/php_curl_http_client_requests/test_php_curl_post_fields_array_multipart
// origin: languages/php/tests/php/test_php_curl_http_client_requests.rs
// vybe-test-mode: compile

$ch = curl_init("https://httpbin.org/post");
curl_setopt($ch, CURLOPT_POST, true);
curl_setopt($ch, CURLOPT_POSTFIELDS, [
    "username" => "alice",
    "file" => "data_content"
]);
curl_close($ch);
