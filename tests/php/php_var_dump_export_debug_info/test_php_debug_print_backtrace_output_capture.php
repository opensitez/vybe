<?php
// vybe-test: php/php_var_dump_export_debug_info/test_php_debug_print_backtrace_output_capture
// origin: languages/php/tests/php/test_php_var_dump_export_debug_info.rs
// vybe-test-mode: compile

function traceTest() {
    ob_start();
    debug_print_backtrace();
    $out = ob_get_clean();
    echo str_contains($out, "traceTest") ? "TRACE_PRINT_OK" : "FAIL";
}
traceTest();
