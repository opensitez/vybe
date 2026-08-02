<?php
// vybe-test: php/file_functions/unlink_file
// origin: languages/php/tests/php/test_file_functions.rs
// vybe-test-mode: compile

$path = '/tmp/vybe_unlink_' . getmypid() . '.txt';
file_put_contents($path, 'tmp');
$ok = unlink($path);
echo $ok ? 'deleted' : 'failed';
