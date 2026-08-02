<?php
// vybe-test: php/php_array_diff_intersect_associations/test_php_array_udiff_value_callback
// origin: languages/php/tests/php/test_php_array_diff_intersect_associations.rs
// vybe-test-mode: compile

$a1 = [1.5, 2.5, 3.5];
$a2 = [1.0, 2.0];
$diff = array_udiff($a1, $a2, fn($a, $b) => (int)$a <=> (int)$b);
echo count($diff) === 1 && $diff[2] == 3.5 ? "UDIFF_OK" : "FAIL";
