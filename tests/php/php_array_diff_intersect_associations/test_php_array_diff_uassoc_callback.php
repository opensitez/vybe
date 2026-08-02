<?php
// vybe-test: php/php_array_diff_intersect_associations/test_php_array_diff_uassoc_callback
// origin: languages/php/tests/php/test_php_array_diff_intersect_associations.rs
// vybe-test-mode: compile

$a1 = ["a" => "green", "b" => "brown"];
$a2 = ["A" => "green", "b" => "yellow"];
$diff = array_diff_uassoc($a1, $a2, fn($a, $b) => strcasecmp($a, $b));
echo count($diff) === 1 && isset($diff["b"]) ? "DIFF_UASSOC_OK" : "FAIL";
