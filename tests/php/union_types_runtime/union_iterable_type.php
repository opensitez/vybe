<?php
// vybe-test: php/union_types_runtime/union_iterable_type
// origin: languages/php/tests/php/test_union_types_runtime.rs

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

function sum(iterable $it): int { $s = 0; foreach ($it as $n) $s += $n; return $s; }
echo sum([1, 2]);

__vybe_check(ob_get_clean(), "3");
