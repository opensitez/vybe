<?php
// vybe-test: php/magic_constants/magic_function_nested
// origin: languages/php/tests/php/test_magic_constants.rs
// vybe-test-mode: compile

function outer(): string {
    function inner(): string { return __FUNCTION__; }
    return __FUNCTION__ . ':' . inner();
}
echo outer();
