<?php
// vybe-test: php/advanced_closures/call_user_func_array_with_closure
// origin: languages/php/tests/php/test_advanced_closures.rs
// vybe-test-mode: compile

$sum = function(int $a, int $b, int $c): int { return $a + $b + $c; };
echo call_user_func_array($sum, [1, 2, 3]);
