<?php
// vybe-test: php/php_stream_socket_pair_bidirectional/test_php_stream_socket_pair_invalid_domain_returns_false
// origin: languages/php/tests/php/test_php_stream_socket_pair_bidirectional.rs
// vybe-test-mode: compile

$res = @stream_socket_pair(99999, STREAM_SOCK_STREAM, 0);
echo $res === false ? "INVALID_DOMAIN_FALSE" : "FAIL";
