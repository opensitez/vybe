<?php
// vybe-test: php/php_stream_socket_pair_bidirectional/test_php_stream_socket_pair_metadata_type
// origin: languages/php/tests/php/test_php_stream_socket_pair_bidirectional.rs
// vybe-test-mode: compile

$pair = stream_socket_pair(STREAM_PF_UNIX, STREAM_SOCK_STREAM, 0);
$meta = stream_get_meta_data($pair[0]);
fclose($pair[0]);
fclose($pair[1]);
echo isset($meta["stream_type"]) ? "PAIR_META_TYPE_OK" : "FAIL";
