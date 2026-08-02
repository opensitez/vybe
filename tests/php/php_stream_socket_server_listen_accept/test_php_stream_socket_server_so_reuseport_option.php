<?php
// vybe-test: php/php_stream_socket_server_listen_accept/test_php_stream_socket_server_so_reuseport_option
// origin: languages/php/tests/php/test_php_stream_socket_server_listen_accept.rs
// vybe-test-mode: compile

$context = stream_context_create(["socket" => ["so_reuseport" => true]]);
$server = @stream_socket_server("tcp://127.0.0.1:0", $errno, $errstr, STREAM_SERVER_BIND | STREAM_SERVER_LISTEN, $context);
if ($server) fclose($server);
echo "REUSEPORT_OPTION_OK";
