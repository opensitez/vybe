<?php
// vybe-test: php/functions/closure_basic
// origin: languages/php/tests/php/test_functions.rs
// vybe-test-mode: compile

$fn = function($x) { return $x * 2; }; echo $fn(5);
