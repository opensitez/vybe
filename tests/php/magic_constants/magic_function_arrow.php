<?php
// vybe-test: php/magic_constants/magic_function_arrow
// origin: languages/php/tests/php/test_magic_constants.rs
// vybe-test-mode: compile

$fn = fn() => __FUNCTION__;
$result = $fn();
echo is_string($result) ? 'string' : 'fail';
