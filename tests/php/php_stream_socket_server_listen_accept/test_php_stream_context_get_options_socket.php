<?php
// vybe-test: php/php_stream_socket_server_listen_accept/test_php_stream_context_get_options_socket
// origin: languages/php/tests/php/test_php_stream_socket_server_listen_accept.rs
// vybe-test-mode: compile

$context = stream_context_create(["socket" => ["bindto" => "127.0.0.1:0"]]);
$opts = stream_context_get_options($context);
echo isset($opts["socket"]["bindto"]) ? "CONTEXT_OPTIONS_GET_OK" : "FAIL";
