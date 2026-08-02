<?php
// vybe-test: php/php_array_diff_intersect_associations/test_php_array_diff_key_compares_keys
// origin: languages/php/tests/php/test_php_array_diff_intersect_associations.rs
// vybe-test-mode: compile

$a1 = [10 => "val1", 20 => "val2", 30 => "val3"];
$a2 = [10 => "different", 40 => "val4"];
$diff = array_diff_key($a1, $a2);
echo count($diff) === 2 && isset($diff[20]) && isset($diff[30]) ? "DIFF_KEY_OK" : "FAIL";
