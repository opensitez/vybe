<?php
// vybe-test: php/memory_get_usage_allocations/memory_get_usage_real
// origin: languages/php/tests/php/test_memory_get_usage_allocations.rs

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

$mem = memory_get_usage(true);
echo is_int($mem) && $mem > 0 ? "ok" : "fail";

__vybe_check(ob_get_clean(), "ok");
