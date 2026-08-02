<?php
// vybe-test: php/array_builtins_extended/array_walk_recursive_leaf_callback
// origin: languages/php/tests/php/test_array_builtins_extended.rs
// vybe-test-mode: compile

$nested = [1, [2, 3], [[4], 5]];
$collected = [];
array_walk_recursive($nested, function($val) use (&$collected) {
    $collected[] = $val * 2;
});
echo implode(",", $collected);
