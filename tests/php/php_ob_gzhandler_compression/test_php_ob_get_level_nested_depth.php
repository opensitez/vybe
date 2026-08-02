<?php
// vybe-test: php/php_ob_gzhandler_compression/test_php_ob_get_level_nested_depth
// origin: languages/php/tests/php/test_php_ob_gzhandler_compression.rs
// vybe-test-mode: compile

$l0 = ob_get_level();
ob_start();
$l1 = ob_get_level();
ob_start();
$l2 = ob_get_level();
ob_end_clean();
ob_end_clean();
echo $l1 === $l0 + 1 && $l2 === $l0 + 2 ? "NESTED_LEVELS_OK" : "FAIL";
