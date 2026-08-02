<?php
// vybe-test: php/php_stream_socket_client_timeout_options/test_php_stream_set_read_buffer_chunk_size
// origin: languages/php/tests/php/test_php_stream_socket_client_timeout_options.rs
// vybe-test-mode: compile

$fp = fopen("php://memory", "r+");
$res = stream_set_read_buffer($fp, 4096);
fclose($fp);
echo is_int($res) ? "READ_BUFFER_OK" : "FAIL";
