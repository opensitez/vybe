<?php
// vybe-test: php/error_reporting_runtime/debug_print_backtrace_captures
// origin: languages/php/tests/php/test_error_reporting_runtime.rs

ob_start();
debug_print_backtrace(DEBUG_BACKTRACE_IGNORE_ARGS, 1);
$out = ob_get_clean();
echo strlen($out) > 0 ? 'trace' : 'empty';
