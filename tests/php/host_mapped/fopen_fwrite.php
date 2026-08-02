<?php
// vybe-test: php/host_mapped/fopen_fwrite
// origin: languages/php/tests/php/test_host_mapped.rs
// vybe-test-mode: compile

$fp = fsockopen('localhost', 80);
fwrite($fp, "GET / HTTP/1.0\r\n\r\n");
$line = fgets($fp);
fclose($fp);
