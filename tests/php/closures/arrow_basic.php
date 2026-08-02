<?php
// vybe-test: php/closures/arrow_basic
// origin: languages/php/tests/php/test_closures.rs
// vybe-test-mode: compile

$fn = fn($x) => $x * 2; echo $fn(5);
