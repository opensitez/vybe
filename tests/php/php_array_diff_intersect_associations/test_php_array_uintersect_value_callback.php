<?php
// vybe-test: php/php_array_diff_intersect_associations/test_php_array_uintersect_value_callback
// origin: languages/php/tests/php/test_php_array_diff_intersect_associations.rs
// vybe-test-mode: compile

$a1 = ["Apple", "banana"];
$a2 = ["apple", "BANANA"];
$inter = array_uintersect($a1, $a2, fn($a, $b) => strcasecmp($a, $b));
echo count($inter) === 2 ? "UINTERSECT_OK" : "FAIL";
