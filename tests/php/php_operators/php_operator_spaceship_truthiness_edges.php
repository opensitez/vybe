<?php
// vybe-test: php/php_operators/php_operator_spaceship_truthiness_edges
// origin: languages/php/tests/php/test_php_operators.rs

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

echo (4 <=> 4), "\n";
echo (-1 <=> 2), "\n";
echo (2 <=> -1), "\n";
echo (true <=> false), "\n";
echo (false <=> true), "\n";
echo (null <=> null), "\n";
echo (true ?: 'fallback'), "\n";
echo ((true && false) ?: 'fallback'), "\n";
echo ((0 ?: 1) <=> (1 ?: 0)), "\n";


__vybe_check(ob_get_clean(), "0\n-1\n1\n1\n-1\n0\n1\nfallback\n0");
