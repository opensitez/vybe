<?php
// vybe-test: php/php_file_path_resolution_canonicalization/test_php_realpath_cache_size_and_get
// origin: languages/php/tests/php/test_php_file_path_resolution_canonicalization.rs
// vybe-test-mode: compile

$cacheSize = realpath_cache_size();
$entries = realpath_cache_get();
echo "Size=$cacheSize Entries=" . (is_array($entries) ? count($entries) : 0);
