<?php
// vybe-test: php/named_args_extended/named_arg_in_match_arm_function
// origin: languages/php/tests/php/test_named_args_extended.rs

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

function format(float $n, int $decimals = 2, string $separator = '.'): string {
    return number_format($n, $decimals, $separator, ',');
}
echo format(n: 1234.5678, decimals: 1);

__vybe_check(ob_get_clean(), "1,234.6");
