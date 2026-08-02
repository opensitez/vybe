<?php
// vybe-test: php/closures/iife_with_args
// origin: languages/php/tests/php/test_closures.rs
// vybe-test-mode: compile

$result = (function($a, $b) { return $a + $b; })(3, 4); echo $result;
