<?php
// vybe-test: php/php_stream_socket_client_timeout_options/test_php_stream_get_meta_data_stream_type
// origin: languages/php/tests/php/test_php_stream_socket_client_timeout_options.rs
// vybe-test-mode: compile

$fp = fopen("php://memory", "r+");
$meta = stream_get_meta_data($fp);
fclose($fp);
echo $meta["stream_type"] === "MEMORY" || str_contains($meta["wrapper_type"], "php") ? "META_STREAM_TYPE_OK" : "FAIL";
