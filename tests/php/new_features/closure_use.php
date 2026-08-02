<?php
// vybe-test: php/new_features/closure_use
// origin: languages/php/tests/php/test_new_features.rs
// vybe-test-mode: compile

$x = 10; $fn = function() use ($x) { return $x; }; echo $fn();
