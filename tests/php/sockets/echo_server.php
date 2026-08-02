<?php
// vybe-test: php/sockets/echo_server
// origin: languages/php/tests/php/test_sockets.rs
// vybe-test-mode: compile

function startServer($port) {
    $server = stream_socket_server('tcp://0.0.0.0:' . $port);
    echo "Listening on port $port\n";
    $client = stream_socket_accept($server);
    $data = stream_get_contents($client);
    socket_write($client, $data);
    socket_close($client);
}
