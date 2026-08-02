<?php
// vybe-test: php/sockets/http_client
// origin: languages/php/tests/php/test_sockets.rs
// vybe-test-mode: compile

$fp = fsockopen('httpbin.org', 80);
fwrite($fp, "GET /get HTTP/1.1\r\nHost: httpbin.org\r\nConnection: close\r\n\r\n");
$response = '';
$line = fgets($fp);
fclose($fp);
echo $response;
