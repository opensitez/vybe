<?php
// vybe-test: php/php_stream_socket_client_timeout_options/test_php_stream_socket_shutdown_read_write
// origin: languages/php/tests/php/test_php_stream_socket_client_timeout_options.rs
// vybe-test-mode: compile

$pair = stream_socket_pair(STREAM_PF_UNIX, STREAM_SOCK_STREAM, 0);
if ($pair) {
    stream_socket_shutdown($pair[0], STREAM_SHUT_WR);
    fclose($pair[0]);
    fclose($pair[1]);
    echo "SHUTDOWN_OK";
}
