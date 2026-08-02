<?php
// vybe-test: php/closures/closure_with_body
// origin: languages/php/tests/php/test_closures.rs
// vybe-test-mode: compile

$fn = function($n) { $result = 1; for ($i = 1; $i <= $n; $i++) { $result *= $i; } return $result; }; echo $fn(5);
