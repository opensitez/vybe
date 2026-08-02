<?php
// vybe-test: php/php_output_buffering_ob_start_clean/test_php_nested_output_buffering_levels
// origin: languages/php/tests/php/test_php_output_buffering_ob_start_clean.rs

ob_start(); // Level 1
echo "Level 1";
ob_start(); // Level 2
echo "Level 2";
$l2 = ob_get_clean();
$l1 = ob_get_clean();
echo "L1=$l1 | L2=$l2";
