<?php
// vybe-test: php/php_stream_socket_pair_bidirectional/test_php_stream_socket_pair_unclosed_buffer_flush
// origin: languages/php/tests/php/test_php_stream_socket_pair_bidirectional.rs
// vybe-test-mode: compile

$pair = stream_socket_pair(STREAM_PF_UNIX, STREAM_SOCK_STREAM, 0);
fwrite($pair[0], "line1\n");
fwrite($pair[0], "line2\n");
$l1 = fgets($pair[1]);
$l2 = fgets($pair[1]);
fclose($pair[0]);
fclose($pair[1]);
echo trim($l1) === "line1" && trim($l2) === "line2" ? "FGETS_PAIR_OK" : "FAIL";
