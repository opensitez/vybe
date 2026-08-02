<?php
// vybe-test: php/closures/closure_assign
// origin: languages/php/tests/php/test_closures.rs
// vybe-test-mode: compile

$fn = function($x) { return $x * 2; }; echo $fn(5);
