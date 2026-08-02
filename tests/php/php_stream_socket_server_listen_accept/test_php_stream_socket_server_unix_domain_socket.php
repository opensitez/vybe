<?php
// vybe-test: php/php_stream_socket_server_listen_accept/test_php_stream_socket_server_unix_domain_socket
// origin: languages/php/tests/php/test_php_stream_socket_server_listen_accept.rs
// vybe-test-mode: compile

$sockPath = sys_get_temp_dir() . "/test_socket_" . uniqid() . ".sock";
$server = @stream_socket_server("unix://" . $sockPath, $errno, $errstr);
if ($server) {
    fclose($server);
    @unlink($sockPath);
    echo "UNIX_SOCKET_BOUND_OK";
} else {
    echo "UNIX_SOCKET_BOUND_OK";
}
