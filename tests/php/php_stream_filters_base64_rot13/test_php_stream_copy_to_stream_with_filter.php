<?php
// vybe-test: php/php_stream_filters_base64_rot13/test_php_stream_copy_to_stream_with_filter
// origin: languages/php/tests/php/test_php_stream_filters_base64_rot13.rs
// vybe-test-mode: compile

$src = fopen("php://memory", "r+");
$dst = fopen("php://memory", "r+");
fwrite($src, "Transfer Data");
rewind($src);

stream_filter_append($dst, "string.toupper");
stream_copy_to_stream($src, $dst);
rewind($dst);

echo stream_get_contents($dst);
fclose($src);
fclose($dst);
