<?php
// vybe-test: php/php_stream_socket_client_timeout_options/test_php_stream_socket_client_async_flags
// origin: languages/php/tests/php/test_php_stream_socket_client_timeout_options.rs
// vybe-test-mode: compile

$fp = @stream_socket_client("tcp://127.0.0.1:65533", $errno, $errstr, 0.5, STREAM_CLIENT_ASYNC_CONNECT);
if ($fp) fclose($fp);
echo "ASYNC_FLAG_OK";
