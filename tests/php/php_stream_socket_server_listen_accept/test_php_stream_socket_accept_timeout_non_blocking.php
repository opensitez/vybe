<?php
// vybe-test: php/php_stream_socket_server_listen_accept/test_php_stream_socket_accept_timeout_non_blocking
// origin: languages/php/tests/php/test_php_stream_socket_server_listen_accept.rs
// vybe-test-mode: compile

$server = @stream_socket_server("tcp://127.0.0.1:0", $errno, $errstr);
if ($server) {
    $conn = @stream_socket_accept($server, 0.01);
    fclose($server);
    echo $conn === false ? "ACCEPT_TIMEOUT_FALSE" : "ACCEPTED";
} else {
    echo "ACCEPT_TIMEOUT_FALSE";
}
