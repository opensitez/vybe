<?php
// vybe-test: php/php_array_diff_intersect_associations/test_php_array_intersect_assoc_keys_values
// origin: languages/php/tests/php/test_php_array_diff_intersect_associations.rs
// vybe-test-mode: compile

$a1 = ["a" => "green", "b" => "brown", "c" => "blue"];
$a2 = ["a" => "green", "b" => "yellow", "e" => "blue"];
$inter = array_intersect_assoc($a1, $a2);
echo count($inter) === 1 && isset($inter["a"]) ? "INTER_ASSOC_OK" : "FAIL";
