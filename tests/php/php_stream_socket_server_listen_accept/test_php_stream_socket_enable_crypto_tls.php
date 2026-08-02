<?php
// vybe-test: php/php_stream_socket_server_listen_accept/test_php_stream_socket_enable_crypto_tls
// origin: languages/php/tests/php/test_php_stream_socket_server_listen_accept.rs
// vybe-test-mode: compile

$fp = fopen("php://memory", "r+");
$res = @stream_socket_enable_crypto($fp, true, STREAM_CRYPTO_METHOD_TLS_CLIENT);
fclose($fp);
echo $res === false ? "ENABLE_CRYPTO_HANDLED" : "FAIL";
