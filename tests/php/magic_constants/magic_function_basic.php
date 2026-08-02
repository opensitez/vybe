<?php
// vybe-test: php/magic_constants/magic_function_basic
// origin: languages/php/tests/php/test_magic_constants.rs
// vybe-test-mode: compile

function myFunc(): string { return __FUNCTION__; }
echo myFunc();
