<?php
// vybe-test: php/closures/use_multiple
// origin: languages/php/tests/php/test_closures.rs
// vybe-test-mode: compile

$a = 1; $b = 2; $fn = function() use ($a, $b) { return $a + $b; }; echo $fn();
