<?php
// vybe-test: php/closures/closure_multi_args
// origin: languages/php/tests/php/test_closures.rs
// vybe-test-mode: compile

$fn = function($a, $b, $c) { return $a + $b + $c; }; echo $fn(1,2,3);
