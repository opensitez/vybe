<?php
// vybe-test: php/sockets/fsockopen_basic
// origin: languages/php/tests/php/test_sockets.rs
// vybe-test-mode: compile

$fp = fsockopen('localhost', 80);
fwrite($fp, "GET / HTTP/1.0\r\n\r\n");
$response = fgets($fp);
fclose($fp);
