<?php
// vybe-test: php/debug_print_backtrace_limit/debug_print_backtrace_limit
// origin: languages/php/tests/php/test_debug_print_backtrace_limit.rs

function a() { b(); }
function b() { c(); }
function c() {
    ob_start();
    debug_print_backtrace(DEBUG_BACKTRACE_IGNORE_ARGS, 2);
    $trace = ob_get_clean();
    echo substr_count($trace, '#') === 2 ? "ok" : "fail";
}
a();
