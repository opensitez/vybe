<?php
// vybe-test: php/file_functions/copy_file
// origin: languages/php/tests/php/test_file_functions.rs
// vybe-test-mode: compile

$src = '/tmp/vybe_copy_src_' . getmypid() . '.txt';
$dst = '/tmp/vybe_copy_dst_' . getmypid() . '.txt';
file_put_contents($src, 'data');
$ok = copy($src, $dst);
echo $ok ? 'copied' : 'failed';
if (file_exists($src)) unlink($src);
if (file_exists($dst)) unlink($dst);
