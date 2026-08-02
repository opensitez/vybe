<?php
// vybe-test: php/file_functions/file_get_contents_with_context
// origin: languages/php/tests/php/test_file_functions.rs
// vybe-test-mode: compile

$opts = ['http' => ['method' => 'GET', 'header' => 'Accept: text/html']];
$ctx = stream_context_create($opts);
$result = file_get_contents('http://example.com', false, $ctx);
echo is_string($result) || $result === false ? 'ok' : 'fail';
