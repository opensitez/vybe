<?php
// vybe-test: php/new_features/closure_use_multiple
// origin: languages/php/tests/php/test_new_features.rs
// vybe-test-mode: compile

$a = 1; $b = 2; $fn = function() use ($a, $b) { return $a + $b; }; echo $fn();
