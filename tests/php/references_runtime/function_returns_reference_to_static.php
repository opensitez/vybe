<?php
// vybe-test: php/references_runtime/function_returns_reference_to_static
// origin: languages/php/tests/php/test_references_runtime.rs

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

function &counter(): int {
    static $n = 0;
    $n++;
    return $n;
}
counter();
echo counter();

__vybe_check(ob_get_clean(), "2");
