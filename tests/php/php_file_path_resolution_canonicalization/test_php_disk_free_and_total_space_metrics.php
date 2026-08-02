<?php
// vybe-test: php/php_file_path_resolution_canonicalization/test_php_disk_free_and_total_space_metrics
// origin: languages/php/tests/php/test_php_file_path_resolution_canonicalization.rs
// vybe-test-mode: compile

$free = disk_free_space(".");
$total = disk_total_space(".");
echo ($free !== false && $total !== false && $total >= $free) ? "SPACE_METRICS_OK" : "FAIL";
