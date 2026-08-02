<?php
// vybe-test: php/closures/arrow_no_params
// origin: languages/php/tests/php/test_closures.rs
// vybe-test-mode: compile

$fn = fn() => 42; echo $fn();
