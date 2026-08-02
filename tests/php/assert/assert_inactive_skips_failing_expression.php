<?php
// vybe-test: php/assert/assert_inactive_skips_failing_expression
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

$prev = assert_options(ASSERT_ACTIVE, 0);
assert(false);
assert_options(ASSERT_ACTIVE, $prev);
echo 'skipped';

__vybe_check(ob_get_clean(), "skipped");
