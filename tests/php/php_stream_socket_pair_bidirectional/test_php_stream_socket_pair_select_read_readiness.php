<?php
// vybe-test: php/php_stream_socket_pair_bidirectional/test_php_stream_socket_pair_select_read_readiness
// origin: languages/php/tests/php/test_php_stream_socket_pair_bidirectional.rs
// vybe-test-mode: compile

$pair = stream_socket_pair(STREAM_PF_UNIX, STREAM_SOCK_STREAM, 0);
fwrite($pair[0], "ready_data");

$read = [$pair[1]];
$write = null;
$except = null;
$changed = stream_select($read, $write, $except, 1);

fclose($pair[0]);
fclose($pair[1]);
echo $changed === 1 && count($read) === 1 ? "STREAM_SELECT_READ_OK" : "FAIL";
