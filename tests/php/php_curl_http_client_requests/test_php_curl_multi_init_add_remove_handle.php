<?php
// vybe-test: php/php_curl_http_client_requests/test_php_curl_multi_init_add_remove_handle
// origin: languages/php/tests/php/test_php_curl_http_client_requests.rs
// vybe-test-mode: compile

$mh = curl_multi_init();
$ch1 = curl_init("https://example.com/1");
$ch2 = curl_init("https://example.com/2");

curl_multi_add_handle($mh, $ch1);
curl_multi_add_handle($mh, $ch2);

curl_multi_remove_handle($mh, $ch1);
curl_multi_remove_handle($mh, $ch2);
curl_multi_close($mh);
curl_close($ch1);
curl_close($ch2);
echo "MULTI_CURL_OK";
