<?php
// vybe-test: php/php_curl_http_client_requests/test_php_curl_escape_unescape_urls
// origin: languages/php/tests/php/test_php_curl_http_client_requests.rs
// vybe-test-mode: compile

$ch = curl_init();
$escaped = curl_escape($ch, "hello world & php");
$unescaped = curl_unescape($ch, $escaped);
curl_close($ch);
echo $unescaped === "hello world & php" ? "ESCAPE_ROUNDTRIP_OK" : "FAIL";
