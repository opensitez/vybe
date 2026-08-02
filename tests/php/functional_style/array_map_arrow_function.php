<?php
// vybe-test: php/functional_style/array_map_arrow_function
// origin: languages/php/tests/php/test_functional_style.rs
// vybe-test-mode: compile

$nums = [1, 2, 3, 4, 5];
$doubled = array_map(fn($n) => $n * 2, $nums);
echo implode(',', $doubled);
