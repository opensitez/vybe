<?php
// vybe-test: php/php_stream_socket_client_timeout_options/test_php_stream_supports_lock_memory_stream
// origin: languages/php/tests/php/test_php_stream_socket_client_timeout_options.rs
// vybe-test-mode: compile

$fp = fopen("php://memory", "r+");
$meta = stream_get_meta_data($fp);
fclose($fp);
echo isset($meta["seekable"]) ? "SEEKABLE_META_OK" : "FAIL";
