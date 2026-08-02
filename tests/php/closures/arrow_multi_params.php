<?php
// vybe-test: php/closures/arrow_multi_params
// origin: languages/php/tests/php/test_closures.rs
// vybe-test-mode: compile

$fn = fn($a, $b) => $a + $b; echo $fn(3, 4);
