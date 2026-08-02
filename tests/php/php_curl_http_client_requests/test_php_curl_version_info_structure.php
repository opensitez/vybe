<?php
// vybe-test: php/php_curl_http_client_requests/test_php_curl_version_info_structure
// origin: languages/php/tests/php/test_php_curl_http_client_requests.rs
// vybe-test-mode: compile

$v = curl_version();
echo "cURL Version: " . $v["version"] . " SSL: " . $v["ssl_version"];
