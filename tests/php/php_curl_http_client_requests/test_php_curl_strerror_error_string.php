<?php
// vybe-test: php/php_curl_http_client_requests/test_php_curl_strerror_error_string
// origin: languages/php/tests/php/test_php_curl_http_client_requests.rs
// vybe-test-mode: compile

$errStr = curl_strerror(CURLE_COULDNT_RESOLVE_HOST);
echo is_string($errStr) && strlen($errStr) > 0 ? "STRERROR_OK" : "FAIL";
