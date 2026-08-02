<?php
// vybe-test: php/functions/arrow_fn
// origin: languages/php/tests/php/test_functions.rs
// vybe-test-mode: compile

$fn = fn($x) => $x * 2; echo $fn(5);
