<?php
// vybe-test: php/php_stream_socket_pair_bidirectional/test_php_stream_socket_pair_large_payload_transfer
// origin: languages/php/tests/php/test_php_stream_socket_pair_bidirectional.rs
// vybe-test-mode: compile

$pair = stream_socket_pair(STREAM_PF_UNIX, STREAM_SOCK_STREAM, 0);
$payload = str_repeat("ABCDEFGH", 1024); // 8KB payload
fwrite($pair[0], $payload);
$received = stream_get_contents($pair[1], strlen($payload));
fclose($pair[0]);
fclose($pair[1]);
echo $received === $payload ? "8KB_PAIR_TRANSFER_OK" : "FAIL";
