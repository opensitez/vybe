<?php
// vybe-test: php/php_stream_socket_client_timeout_options/test_php_stream_set_write_buffer_unbuffered
// origin: languages/php/tests/php/test_php_stream_socket_client_timeout_options.rs
// vybe-test-mode: compile

$fp = fopen("php://memory", "r+");
$res = stream_set_write_buffer($fp, 0);
fclose($fp);
echo is_int($res) ? "WRITE_BUFFER_UNBUFFERED_OK" : "FAIL";
