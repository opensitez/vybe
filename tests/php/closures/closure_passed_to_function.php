<?php
// vybe-test: php/closures/closure_passed_to_function
// origin: languages/php/tests/php/test_closures.rs
// vybe-test-mode: compile

function apply($fn, $val) { return $fn($val); } echo apply(fn($x) => $x + 1, 41);
