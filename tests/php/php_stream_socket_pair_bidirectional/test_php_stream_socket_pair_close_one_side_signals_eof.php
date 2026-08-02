<?php
// vybe-test: php/php_stream_socket_pair_bidirectional/test_php_stream_socket_pair_close_one_side_signals_eof
// origin: languages/php/tests/php/test_php_stream_socket_pair_bidirectional.rs
// vybe-test-mode: compile

$pair = stream_socket_pair(STREAM_PF_UNIX, STREAM_SOCK_STREAM, 0);
fclose($pair[0]);
$data = fread($pair[1], 10);
$eof = feof($pair[1]);
fclose($pair[1]);
echo $data === "" && $eof ? "CLOSED_SIDE_EOF_OK" : "FAIL";
