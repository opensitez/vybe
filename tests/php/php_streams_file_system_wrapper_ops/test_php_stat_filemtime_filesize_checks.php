<?php
// vybe-test: php/php_streams_file_system_wrapper_ops/test_php_stat_filemtime_filesize_checks
// origin: languages/php/tests/php/test_php_streams_file_system_wrapper_ops.rs
// vybe-test-mode: compile

$tmp = tempnam(sys_get_temp_dir(), "stat_test_");
file_put_contents($tmp, "12345");
echo filesize($tmp) . " bytes mtime=" . filemtime($tmp);
unlink($tmp);
