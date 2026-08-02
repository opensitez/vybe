<?php
// vybe-test: php/array_extra_builtins/array_udiff_user_value_comparison
// origin: languages/php/tests/php/test_array_extra_builtins.rs
// vybe-test-mode: compile

$a = [1, 2, 3, 4, 5];
$b = [3, 4, 5, 6, 7];
$diff = array_udiff($a, $b, function($x, $y) { return $x - $y; });
echo implode(",", $diff);
