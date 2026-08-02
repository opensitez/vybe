<?php
// vybe-test: php/named_arguments/named_args_combined_with_spread
// origin: languages/php/tests/php/test_named_arguments.rs

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

function sumThree(int $a, int $b, int $c): int { return $a + $b + $c; }
$extra = [2, 3];
echo sumThree(a: 1, ...$extra) . "\n";

__vybe_check(ob_get_clean(), "6");
