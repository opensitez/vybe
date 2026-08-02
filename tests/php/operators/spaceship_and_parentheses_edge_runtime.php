<?php
// vybe-test: php/operators/spaceship_and_parentheses_edge_runtime
// origin: languages/php/tests/php/test_operators.rs

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

echo (1 <=> 2) . '|';
echo (2 <=> 1) . '|';
echo (2 <=> 2) . '|';
echo (3 + 5 <=> 4 + 1) . '|';
echo (false <=> true) . '|';
echo ((5 < 3) <=> (2 < 4));

__vybe_check(ob_get_clean(), "-1|1|0|1|-1|-1");
