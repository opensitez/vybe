<?php
// vybe-test: php/function_builtins/array_map_null_callback_zip
// origin: languages/php/tests/php/test_function_builtins.rs
// vybe-test-mode: compile

$a = [1, 2, 3];
$b = ['a', 'b', 'c'];
$zipped = array_map(null, $a, $b);
echo count($zipped);
echo $zipped[0][0] . $zipped[0][1];
