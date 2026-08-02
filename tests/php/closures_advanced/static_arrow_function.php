<?php
// vybe-test: php/closures_advanced/static_arrow_function
// origin: languages/php/tests/php/test_closures_advanced.rs
// vybe-test-mode: compile

$fn = static fn(int $a, int $b) => $a + $b;
echo $fn(10, 32);
