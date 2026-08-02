<?php
// vybe-test: php/php_stream_filters_base64_rot13/test_php_stream_filter_prepend_order
// origin: languages/php/tests/php/test_php_stream_filters_base64_rot13.rs
// vybe-test-mode: compile

$stream = fopen("php://memory", "r+");
stream_filter_append($stream, "string.rot13");
stream_filter_prepend($stream, "string.toupper");
fwrite($stream, "abc");
rewind($stream);
echo stream_get_contents($stream);
fclose($stream);
