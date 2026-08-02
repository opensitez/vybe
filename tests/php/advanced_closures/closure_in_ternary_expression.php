<?php
// vybe-test: php/advanced_closures/closure_in_ternary_expression
// origin: languages/php/tests/php/test_advanced_closures.rs
// vybe-test-mode: compile

$flag = true;
$transform = $flag ? fn(int $x) => $x * 10 : fn(int $x) => $x;
echo $transform(5);
