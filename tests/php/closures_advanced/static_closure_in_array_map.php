<?php
// vybe-test: php/closures_advanced/static_closure_in_array_map
// origin: languages/php/tests/php/test_closures_advanced.rs
// vybe-test-mode: compile

$nums = [1, 2, 3, 4, 5];
$squares = array_map(static fn(int $n) => $n ** 2, $nums);
echo implode(',', $squares);
