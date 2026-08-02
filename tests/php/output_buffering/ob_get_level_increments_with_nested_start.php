<?php
// vybe-test: php/output_buffering/ob_get_level_increments_with_nested_start
// origin: languages/php/tests/php/test_output_buffering.rs

$base = ob_get_level();
ob_start();
$inside = ob_get_level();
ob_end_clean();
echo $inside - $base;
