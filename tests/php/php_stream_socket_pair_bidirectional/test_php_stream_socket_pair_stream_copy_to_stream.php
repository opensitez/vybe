<?php
// vybe-test: php/php_stream_socket_pair_bidirectional/test_php_stream_socket_pair_stream_copy_to_stream
// origin: languages/php/tests/php/test_php_stream_socket_pair_bidirectional.rs
// vybe-test-mode: compile

$pair = stream_socket_pair(STREAM_PF_UNIX, STREAM_SOCK_STREAM, 0);
$mem = fopen("php://memory", "r+");
fwrite($mem, "Stream Copy Test Data");
rewind($mem);

stream_copy_to_stream($mem, $pair[0]);
$copied = stream_get_contents($pair[1], 21);

fclose($mem);
fclose($pair[0]);
fclose($pair[1]);
echo $copied === "Stream Copy Test Data" ? "STREAM_COPY_PAIR_OK" : "FAIL";
