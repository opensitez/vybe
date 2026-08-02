<?php
// vybe-test: php/php_curl_http_client_requests/test_php_curl_reset_handle_options
// origin: languages/php/tests/php/test_php_curl_http_client_requests.rs
// vybe-test-mode: compile

$ch = curl_init("https://example.com");
curl_setopt($ch, CURLOPT_TIMEOUT, 5);
curl_reset($ch);
$url = curl_getinfo($ch, CURLINFO_EFFECTIVE_URL);
curl_close($ch);
echo "Reset OK";
