<?php
// vybe-test: php/output_functions/ob_get_level_nesting
// origin: languages/php/tests/php/test_output_functions.rs
// vybe-test-mode: compile

$l0 = ob_get_level();
ob_start();
$l1 = ob_get_level();
ob_start();
$l2 = ob_get_level();
ob_end_clean();
ob_end_clean();
$l3 = ob_get_level();
echo "$l0,$l1,$l2,$l3";
