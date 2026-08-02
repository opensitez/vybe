<?php
// vybe-test: php/error_handler/handler_does_not_run_when_assertion_disabled
// origin: languages/php/tests/php/test_error_handler.rs

function __vybe_check($got, $want) {
    // Match the Rust harness's normalisation: strip \r, then drop trailing
    // newlines (it split on "\n" and popped empty trailing elements).
    $got = str_replace("\r", "", $got);
    $got = rtrim($got, "\n");
    if ($got !== $want) {
        echo "FAIL: want [" . $want . "] got [" . $got . "]\n";
        throw new Exception("assertion failed");
    }
    // Replay the program's own output so running the file by hand still
    // behaves like the program it was extracted from.
    echo $got;
    if ($got !== "") {
        echo "\n";
    }
}

ob_start();

$n = 0;
set_error_handler(function() use (&$n): bool { $n++; return true; });
$old = assert_options(ASSERT_ACTIVE, 0);
assert(false);
assert_options(ASSERT_ACTIVE, $old);
restore_error_handler();
echo $n;

__vybe_check(ob_get_clean(), "4");
