<?php
// vybe-test: php/array_builtins_extended/array_intersect_common_values
// origin: languages/php/tests/php/test_array_builtins_extended.rs
// vybe-test-mode: compile

$a = [1, 2, 3, 4, 5];
$b = [3, 4, 5, 6, 7];
$c = [4, 5, 8, 9];
$common = array_intersect($a, $b, $c);
echo implode(",", $common);
