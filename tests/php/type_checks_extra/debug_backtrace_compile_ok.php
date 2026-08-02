<?php
// vybe-test: php/type_checks_extra/debug_backtrace_compile_ok
// origin: languages/php/tests/php/test_type_checks_extra.rs
// vybe-test-mode: compile

function inner() {
    $bt = debug_backtrace();
    return count($bt);
}
function outer() {
    return inner();
}
echo outer();
