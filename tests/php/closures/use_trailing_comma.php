<?php
// vybe-test: php/closures/use_trailing_comma
// origin: languages/php/tests/php/test_closures.rs
// vybe-test-mode: compile

$a = 1; $fn = function() use ($a,) { return $a; };
