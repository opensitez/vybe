<?php
// vybe-test: php/match_expressions/match_in_return_statement
// origin: languages/php/tests/php/test_match_expressions.rs

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

function sign(int $n): string {
    return match (true) {
        $n < 0 => 'neg',
        $n > 0 => 'pos',
        default => 'zero',
    };
}
echo sign(0);

__vybe_check(ob_get_clean(), "zero");
