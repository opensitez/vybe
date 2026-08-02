<?php
// vybe-test: php/type_checks_extra/debug_print_backtrace_compile_ok
// origin: languages/php/tests/php/test_type_checks_extra.rs
// vybe-test-mode: compile

function traced() {
    debug_print_backtrace();
}
traced();
