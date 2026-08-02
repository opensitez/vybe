<?php
// vybe-test: php/assert/assert_inside_nested_function_bubbles
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

assert_options(ASSERT_EXCEPTION, 1);
function check(bool $ok): void { assert($ok); }
try { check(false); echo 'no'; }
catch (AssertionError $e) { echo 'nested'; }

__vybe_check(ob_get_clean(), "nested");
