<?php
// vybe-test: php/array_extra_builtins/array_uintersect_user_value_comparison
// origin: languages/php/tests/php/test_array_extra_builtins.rs
// vybe-test-mode: compile

$a = [1, 2, 3, 4];
$b = [2, 3, 5, 6];
$common = array_uintersect($a, $b, function($x, $y) { return $x - $y; });
echo implode(",", $common);
