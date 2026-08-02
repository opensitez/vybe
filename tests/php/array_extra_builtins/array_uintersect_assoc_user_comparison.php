<?php
// vybe-test: php/array_extra_builtins/array_uintersect_assoc_user_comparison
// origin: languages/php/tests/php/test_array_extra_builtins.rs
// vybe-test-mode: compile

$a = ["a" => 1, "b" => 2, "c" => 3];
$b = ["a" => 1, "b" => 9, "c" => 3];
$result = array_uintersect_assoc($a, $b, function($v1, $v2) { return $v1 - $v2; });
echo implode(",", array_keys($result));
