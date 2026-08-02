<?php
// vybe-test: php/magic_constants/magic_function_closure
// origin: languages/php/tests/php/test_magic_constants.rs
// vybe-test-mode: compile

$fn = function(): string { return __FUNCTION__; };
echo $fn();  // "{closure}"
