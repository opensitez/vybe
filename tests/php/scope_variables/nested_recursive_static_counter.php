<?php
// vybe-test: php/scope_variables/nested_recursive_static_counter
// origin: languages/php/tests/php/test_scope_variables.rs

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

function nested_static_counter(int $n): int {
    static $acc = 0;
    $acc += 1;
    if ($n <= 1) { return $acc; }
    return $acc + nested_static_counter($n - 1);
}
echo nested_static_counter(1) . '|' . nested_static_counter(2);

__vybe_check(ob_get_clean(), "1|5");
