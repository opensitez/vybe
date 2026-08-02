<?php
// vybe-test: php/php_stream_socket_client_timeout_options/test_php_stream_socket_sendto_recvfrom
// origin: languages/php/tests/php/test_php_stream_socket_client_timeout_options.rs
// vybe-test-mode: compile

$pair = stream_socket_pair(STREAM_PF_UNIX, STREAM_SOCK_DGRAM, 0);
if ($pair) {
    stream_socket_sendto($pair[0], "UDP Packet Data");
    $data = stream_socket_recvfrom($pair[1], 15);
    fclose($pair[0]);
    fclose($pair[1]);
    echo $data === "UDP Packet Data" ? "SENDTO_RECVFROM_OK" : "FAIL";
}
