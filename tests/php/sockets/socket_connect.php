<?php
// vybe-test: php/sockets/socket_connect
// origin: languages/php/tests/php/test_sockets.rs
// vybe-test-mode: compile

$sock = socket_connect('127.0.0.1', 8080);
socket_write($sock, "Hello server");
$data = socket_read($sock);
socket_close($sock);
