<?php
// vybe-test: php/php_curl_http_client_requests/test_php_curl_copy_handle_duplication
// origin: languages/php/tests/php/test_php_curl_http_client_requests.rs
// vybe-test-mode: compile

$ch = curl_init("https://example.com");
curl_setopt($ch, CURLOPT_USERAGENT, "TestAgent");
$copy = curl_copy_handle($ch);
curl_close($ch);
curl_close($copy);
echo "Copy OK";
