<?php
// vybe-test: php/array_extra_builtins/array_udiff_assoc_user_value_comparison
// origin: languages/php/tests/php/test_array_extra_builtins.rs
// vybe-test-mode: compile

$a = ["x" => 1, "y" => 2, "z" => 3];
$b = ["x" => 1, "y" => 9, "w" => 3];
$diff = array_udiff_assoc($a, $b, function($v1, $v2) { return $v1 - $v2; });
echo implode(",", array_keys($diff));
