<?php
// vybe-test: php/programs/factorial_iterative_seven
// origin: languages/php/tests/php/test_programs.rs

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

function factIter(int $n): int {
    $r = 1;
    for ($i = 2; $i <= $n; $i++) $r *= $i;
    return $r;
}
echo factIter(7) . "\n";

__vybe_check(ob_get_clean(), "5040");
