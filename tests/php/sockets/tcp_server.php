<?php
// vybe-test: php/sockets/tcp_server
// origin: languages/php/tests/php/test_sockets.rs
// vybe-test-mode: compile

$server = stream_socket_server('tcp://0.0.0.0:9000');
$client = stream_socket_accept($server);
$data = stream_get_contents($client);
echo $data;
