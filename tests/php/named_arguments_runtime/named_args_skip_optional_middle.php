<?php
// vybe-test: php/named_arguments_runtime/named_args_skip_optional_middle
// origin: languages/php/tests/php/test_named_arguments_runtime.rs

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

function f(int $a, int $b = 10, int $c = 20): int { return $a + $b + $c; }
echo f(1, c: 5);

__vybe_check(ob_get_clean(), "16");
