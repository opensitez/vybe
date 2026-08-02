<?php
// vybe-test: php/sockets/udp_socket
// origin: languages/php/tests/php/test_sockets.rs
// vybe-test-mode: compile

$sock = socket_create(AF_INET, SOCK_DGRAM, SOL_UDP);
socket_sendto($sock, "hello", 5, 0, '127.0.0.1', 9999);
$data = socket_recvfrom($sock);
