<?php
// vybe-test: php/assert/assert_options_restore_exception_flag
// origin: languages/php/tests/php/test_assert.rs

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

$was = assert_options(ASSERT_EXCEPTION, 1);
$now = assert_options(ASSERT_EXCEPTION);
assert_options(ASSERT_EXCEPTION, $was);
echo $now === 1 ? 'on' : 'off';

__vybe_check(ob_get_clean(), "on");
