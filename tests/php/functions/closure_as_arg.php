<?php
// vybe-test: php/functions/closure_as_arg
// origin: languages/php/tests/php/test_functions.rs
// vybe-test-mode: compile

function apply($fn, $val) { return $fn($val); } echo apply(fn($x) => $x + 1, 41);
