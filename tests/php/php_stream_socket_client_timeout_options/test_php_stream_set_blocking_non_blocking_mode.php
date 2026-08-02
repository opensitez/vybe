<?php
// vybe-test: php/php_stream_socket_client_timeout_options/test_php_stream_set_blocking_non_blocking_mode
// origin: languages/php/tests/php/test_php_stream_socket_client_timeout_options.rs
// vybe-test-mode: compile

$sockets = stream_socket_pair(STREAM_PF_UNIX, STREAM_SOCK_STREAM, STREAM_IPPROTO_IP);
if ($sockets) {
    stream_set_blocking($sockets[0], false);
    $meta = stream_get_meta_data($sockets[0]);
    fclose($sockets[0]);
    fclose($sockets[1]);
    echo !$meta["blocked"] ? "NON_BLOCKING_OK" : "FAIL";
}
