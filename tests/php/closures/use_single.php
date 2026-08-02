<?php
// vybe-test: php/closures/use_single
// origin: languages/php/tests/php/test_closures.rs
// vybe-test-mode: compile

$x = 10; $fn = function() use ($x) { return $x; }; echo $fn();
