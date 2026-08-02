<?php
// vybe-test: php/php_stream_socket_client_timeout_options/test_php_stream_socket_get_name_local_remote
// origin: languages/php/tests/php/test_php_stream_socket_client_timeout_options.rs
// vybe-test-mode: compile

$sockets = stream_socket_pair(STREAM_PF_UNIX, STREAM_SOCK_STREAM, STREAM_IPPROTO_IP);
if ($sockets) {
    $peer = stream_socket_get_name($sockets[0], true);
    $local = stream_socket_get_name($sockets[0], false);
    fclose($sockets[0]);
    fclose($sockets[1]);
    echo "GET_NAME_OK";
}
