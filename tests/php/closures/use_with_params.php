<?php
// vybe-test: php/closures/use_with_params
// origin: languages/php/tests/php/test_closures.rs
// vybe-test-mode: compile

$factor = 3; $fn = function($x) use ($factor) { return $x * $factor; }; echo $fn(5);
