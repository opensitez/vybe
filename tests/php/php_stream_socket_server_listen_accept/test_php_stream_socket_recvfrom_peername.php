<?php
// vybe-test: php/php_stream_socket_server_listen_accept/test_php_stream_socket_recvfrom_peername
// origin: languages/php/tests/php/test_php_stream_socket_server_listen_accept.rs
// vybe-test-mode: compile

$pair = stream_socket_pair(STREAM_PF_UNIX, STREAM_SOCK_DGRAM, 0);
if ($pair) {
    stream_socket_sendto($pair[0], "data");
    $data = stream_socket_recvfrom($pair[1], 10, 0, $peer);
    fclose($pair[0]);
    fclose($pair[1]);
    echo $data === "data" ? "RECVFROM_PEER_OK" : "FAIL";
}
