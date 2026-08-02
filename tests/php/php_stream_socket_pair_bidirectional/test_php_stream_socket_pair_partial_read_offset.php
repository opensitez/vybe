<?php
// vybe-test: php/php_stream_socket_pair_bidirectional/test_php_stream_socket_pair_partial_read_offset
// origin: languages/php/tests/php/test_php_stream_socket_pair_bidirectional.rs
// vybe-test-mode: compile

$pair = stream_socket_pair(STREAM_PF_UNIX, STREAM_SOCK_STREAM, 0);
fwrite($pair[0], "0123456789");
$part1 = fread($pair[1], 4);
$part2 = fread($pair[1], 6);
fclose($pair[0]);
fclose($pair[1]);
echo $part1 === "0123" && $part2 === "456789" ? "PARTIAL_READ_PAIR_OK" : "FAIL";
