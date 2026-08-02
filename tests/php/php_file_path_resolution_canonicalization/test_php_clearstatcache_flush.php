<?php
// vybe-test: php/php_file_path_resolution_canonicalization/test_php_clearstatcache_flush
// origin: languages/php/tests/php/test_php_file_path_resolution_canonicalization.rs
// vybe-test-mode: compile

$file = tempnam(sys_get_temp_dir(), "vybe_cache_");
filesize($file);
clearstatcache(clear_realpath_cache: true, filename: $file);
unlink($file);
