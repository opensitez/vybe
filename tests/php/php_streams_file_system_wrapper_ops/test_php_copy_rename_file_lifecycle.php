<?php
// vybe-test: php/php_streams_file_system_wrapper_ops/test_php_copy_rename_file_lifecycle
// origin: languages/php/tests/php/test_php_streams_file_system_wrapper_ops.rs
// vybe-test-mode: compile

$src = tempnam(sys_get_temp_dir(), "src_");
$dst = sys_get_temp_dir() . "/dst_file.txt";
file_put_contents($src, "copy test");
copy($src, $dst);
echo is_file($dst) ? "COPIED" : "FAIL";
unlink($src);
unlink($dst);
