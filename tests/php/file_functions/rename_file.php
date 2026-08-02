<?php
// vybe-test: php/file_functions/rename_file
// origin: languages/php/tests/php/test_file_functions.rs
// vybe-test-mode: compile

$src = '/tmp/vybe_rename_src_' . getmypid() . '.txt';
$dst = '/tmp/vybe_rename_dst_' . getmypid() . '.txt';
file_put_contents($src, 'data');
$ok = rename($src, $dst);
echo $ok ? 'renamed' : 'failed';
if (file_exists($dst)) unlink($dst);
