<?php
// vybe-test: php/php_stream_filters_base64_rot13/test_php_stream_bucket_new_creation
// origin: languages/php/tests/php/test_php_stream_filters_base64_rot13.rs
// vybe-test-mode: compile

$stream = fopen("php://memory", "r+");
$bucket = stream_bucket_new($stream, "Bucket content");
echo is_object($bucket) ? "BUCKET_CREATED" : "FAIL";
fclose($stream);
