<?php
// vybe-test: php/array_builtins_extended/array_pad_right_and_left
// origin: languages/php/tests/php/test_array_builtins_extended.rs
// vybe-test-mode: compile

$a = [1, 2, 3];
$right = array_pad($a, 6, 0);
echo count($right);
echo $right[5];
$left = array_pad($a, -6, 9);
echo $left[0];
